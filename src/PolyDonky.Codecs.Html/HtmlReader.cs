using System.Globalization;
using System.Text;
using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using PolyDonky.Core;
using PdBlock     = PolyDonky.Core.Block;
using PdTable     = PolyDonky.Core.Table;
using PdTableRow  = PolyDonky.Core.TableRow;
using PdTableCell = PolyDonky.Core.TableCell;

namespace PolyDonky.Codecs.Html;

/// <summary>
/// HTML5 리더 — AngleSharp DOM 파서를 사용해 HTML 문서를 PolyDonkyument 모델로 매핑한다.
///
/// 지원 요소:
///   블록 — h1~h6, p, ul/ol/li (중첩, GFM 작업 목록 input[type=checkbox]),
///         blockquote (중첩 깊이), pre/code (언어 클래스 보존), hr,
///         table/thead/tbody/tfoot/tr/td/th (셀 정렬·헤더 분리),
///         div/section/article/nav/aside/main/header/footer/figure → 자식 평탄화
///         figcaption → 직전 ImageBlock 의 Description 으로 흡수
///         img → ImageBlock (data: URI 는 바이너리 디코드, 그 외는 ResourcePath)
///         video/audio/iframe → 모노스페이스 텍스트 fallback
///   인라인 — strong/b → Bold, em/i → Italic, u → Underline,
///           s/strike/del → Strikethrough, sub/sup → Sub/Super,
///           code/kbd/samp/var → 모노스페이스, mark → 노란 배경,
///           a[href] → Run.Url, br → 줄바꿈, span/font[style] → 인라인 스타일 파싱
///           HTML entity → 디코드, 알 수 없는 인라인 → 텍스트만 추출
///   <head> — &lt;title&gt; 은 첫 H1 이 없으면 무시 (Core 메타데이터에 매핑하지 않음).
///   script/style/noscript/template/svg/math 등은 무시.
/// </summary>
public sealed class HtmlReader : IDocumentReader
{
    public string FormatId => "html";

    /// <summary>
    /// 한 문서가 만들어낼 수 있는 최대 블록 수 — 0 또는 음수면 제한 없음.
    /// 큰 HTML(예: 복잡한 위키 페이지)은 수만 개의 단락을 만들어 WPF FlowDocument 의
    /// 레이아웃·페이지네이션이 분 단위로 멈출 수 있다 (`GetCharacterRect` 가 전체 측정 강제).
    /// 기본 10,000 — 한도 도달 시 잘라내고 마지막에 경고 단락을 추가한다.
    /// </summary>
    public int MaxBlocks { get; init; } = 10_000;

    private static readonly HtmlParser Parser = new();

    public PolyDonkyument Read(Stream input)
    {
        ArgumentNullException.ThrowIfNull(input);
        using var sr = new StreamReader(input, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: true);
        return FromHtml(sr.ReadToEnd(), MaxBlocks);
    }

    public static PolyDonkyument FromHtml(string source) => FromHtml(source, maxBlocks: 10_000);

    public static PolyDonkyument FromHtml(string source, int maxBlocks)
    {
        ArgumentNullException.ThrowIfNull(source);
        var doc      = Parser.ParseDocument(source);
        var pd       = new PolyDonkyument();
        var section  = new Section();
        pd.Sections.Add(section);

        // <body> 가 없는 단편(fragment) 도 안전하게 처리.
        INode root = doc.Body ?? (INode?)doc.DocumentElement ?? doc;

        // 각주/미주 섹션 먼저 추출 후 DOM 에서 제거 (본문 순회 전).
        ExtractNotesSections(root, pd);

        var ctx = new InlineCtx { Shared = new ReadShared { MaxBlocks = maxBlocks } };
        ProcessChildren(root, section.Blocks, ctx);

        if (ctx.Shared.Truncated)
        {
            var warn = new Paragraph();
            warn.AddText($"[잘림: 원본 HTML 의 블록 수가 한도({maxBlocks:N0})를 초과했습니다]",
                new RunStyle { Italic = true });
            section.Blocks.Add(warn);
            // 호출측(메인 앱)이 사용자에게 한도 없이 재시도 안내를 띄울 수 있게 표시.
            pd.Metadata.Custom["html.truncated"]    = "true";
            pd.Metadata.Custom["html.maxBlocks"]    = maxBlocks.ToString();
        }

        // HTML 문서는 일반적으로 본문 블록이 매우 많아 WPF FlowDocument 의 정밀
        // 페이지네이션(블록별 GetCharacterRect)이 브라우저와 다르게 매우 느리다.
        // 메인 앱의 Paginate 가 이 플래그를 보고 fast-path(전체 page 0 배정) 를 쓰게 한다.
        pd.Metadata.Custom["pagination.degraded"] = "true";

        return pd;
    }

    // ── 각주/미주 섹션 추출 ──────────────────────────────────────────────

    private static void ExtractNotesSections(INode root, PolyDonkyument pd)
    {
        // IDocument 또는 IElement 에서 QuerySelectorAll 사용 가능.
        IParentNode? queryRoot = root as IParentNode;
        if (queryRoot is null) return;

        // <section class="footnotes"> 처리.
        foreach (var sect in queryRoot.QuerySelectorAll("section.footnotes").ToList())
        {
            ParseAndRemoveNotesSection(sect, pd.Footnotes, "fn-");
        }

        // <section class="endnotes"> 처리.
        foreach (var sect in queryRoot.QuerySelectorAll("section.endnotes").ToList())
        {
            ParseAndRemoveNotesSection(sect, pd.Endnotes, "en-");
        }
    }

    private static void ParseAndRemoveNotesSection(IElement sect, IList<FootnoteEntry> target, string idPrefix)
    {
        foreach (var li in sect.QuerySelectorAll("li"))
        {
            var liId = li.GetAttribute("id");
            if (string.IsNullOrEmpty(liId)) continue;

            // <li id="fn-N"> 또는 <li id="en-N">
            var entry = new FootnoteEntry { Id = liId };
            // 복귀 링크(<a href="#fnref-N">↩</a>) 제거 후 내용 파싱.
            foreach (var backLink in li.QuerySelectorAll("a[href^=\"#fnref-\"],a[href^=\"#enref-\"]").ToList())
                backLink.Remove();

            var p = new Paragraph();
            AppendInline(p, li);
            if (p.Runs.Count > 0)
                entry.Blocks.Add(p);
            if (entry.Blocks.Count == 0)
                entry.Blocks.Add(new Paragraph());
            target.Add(entry);
        }
        sect.Remove();
    }

    // ── 처리 컨텍스트 ────────────────────────────────────────────────────

    /// <summary>리더 전체에서 공유되는 상태 — 한도/잘림 플래그.</summary>
    private sealed class ReadShared
    {
        public int  MaxBlocks;
        public bool Truncated;
    }

    private sealed class InlineCtx
    {
        public ListMarker?  Marker;
        public int          QuoteLevel;
        public int          ListLevel;
        public ReadShared   Shared = new();
        public bool LimitReached(IList<PdBlock> target)
            => Shared.MaxBlocks > 0 && target.Count >= Shared.MaxBlocks;
    }

    private static InlineCtx With(InlineCtx baseCtx, ListMarker? marker = null, int? quote = null, int? listLvl = null)
        => new()
        {
            Marker     = marker     ?? baseCtx.Marker,
            QuoteLevel = quote      ?? baseCtx.QuoteLevel,
            ListLevel  = listLvl    ?? baseCtx.ListLevel,
            Shared     = baseCtx.Shared,  // 한도/잘림 플래그는 전체 트리에서 공유.
        };

    // ── 블록 처리 ────────────────────────────────────────────────────────

    private static void ProcessChildren(INode parent, IList<PdBlock> target, InlineCtx ctx)
    {
        foreach (var node in parent.ChildNodes)
        {
            if (ctx.LimitReached(target)) { ctx.Shared.Truncated = true; break; }
            ProcessNode(node, target, ctx);
        }
    }

    private static void ProcessNode(INode node, IList<PdBlock> target, InlineCtx ctx)
    {
        if (node is IText txt)
        {
            // 블록 컨텍스트의 텍스트는 단락으로 감싸 수집.
            if (string.IsNullOrWhiteSpace(txt.Data)) return;
            var p = new Paragraph();
            p.Style.QuoteLevel = ctx.QuoteLevel;
            p.Style.ListMarker = CloneMarker(ctx.Marker);
            p.AddText(NormalizeWhitespace(txt.Data));
            target.Add(p);
            return;
        }

        if (node is not IElement el) return;

        switch (el.LocalName)
        {
            case "h1": case "h2": case "h3": case "h4": case "h5": case "h6":
            {
                var p = new Paragraph();
                p.Style.Outline    = (OutlineLevel)(el.LocalName[1] - '0');
                p.Style.QuoteLevel = ctx.QuoteLevel;
                p.Style.ListMarker = CloneMarker(ctx.Marker);
                ApplyBlockStyle(p, el);
                AppendInline(p, el);
                target.Add(p);
                break;
            }

            case "p":
            {
                var p = new Paragraph();
                p.Style.QuoteLevel = ctx.QuoteLevel;
                p.Style.ListMarker = CloneMarker(ctx.Marker);
                ApplyBlockAlignment(p, el);
                AppendInline(p, el);
                if (p.Runs.Count > 0) target.Add(p);
                break;
            }

            case "br":
            {
                // 단독 br — 빈 단락 추가.
                var p = new Paragraph();
                p.Style.QuoteLevel = ctx.QuoteLevel;
                target.Add(p);
                break;
            }

            case "hr":
            {
                var p = new Paragraph();
                p.Style.IsThematicBreak = true;
                p.Style.QuoteLevel      = ctx.QuoteLevel;
                p.Style.ListMarker      = CloneMarker(ctx.Marker);
                target.Add(p);
                break;
            }

            case "blockquote":
            {
                ProcessChildren(el, target, With(ctx, quote: ctx.QuoteLevel + 1));
                break;
            }

            case "ul":
            case "ol":
            {
                ProcessList(el, target, ctx);
                break;
            }

            case "pre":
            {
                target.Add(BuildPreCodeParagraph(el, ctx));
                break;
            }

            case "code":
            {
                // 단독 <code>(블록) — 모노스페이스 단락. <pre> 안의 <code> 는 위에서 처리됨.
                var p = new Paragraph();
                p.Style.QuoteLevel = ctx.QuoteLevel;
                p.Style.ListMarker = CloneMarker(ctx.Marker);
                p.Style.CodeLanguage = ExtractCodeLanguage(el);
                p.AddText(el.TextContent, MonoStyle());
                target.Add(p);
                break;
            }

            case "table":
            {
                target.Add(BuildTable(el, ctx));
                break;
            }

            case "img":
            {
                target.Add(BuildImage(el));
                break;
            }

            case "figure":
            {
                BuildFigure(el, target, ctx);
                break;
            }

            // 시멘틱 sectioning + 일반 컨테이너 — 자식을 그대로 펼친다.
            case "div":  case "section": case "article":
            case "main": case "aside":   case "header":  case "footer":
            case "nav":  case "details": case "summary":
            case "dl":   case "dt":      case "dd":
            case "form": case "fieldset":
            {
                ProcessChildren(el, target, ctx);
                break;
            }

            // 무시할 요소 — script/style/template/noscript/...
            case "script": case "style":  case "template":
            case "noscript": case "svg":   case "math":
            case "head":    case "meta":   case "link": case "title":
            {
                break;
            }

            // <video>, <audio>, <iframe> — 텍스트 fallback (src URL 보존).
            case "video": case "audio": case "iframe":
            case "embed": case "object":
            {
                var p = new Paragraph();
                p.Style.QuoteLevel = ctx.QuoteLevel;
                p.Style.ListMarker = CloneMarker(ctx.Marker);
                var label = $"[{el.LocalName} {el.GetAttribute("src") ?? el.GetAttribute("data") ?? ""}]";
                p.AddText(label, MonoStyle());
                target.Add(p);
                break;
            }

            // <html>, <body> 등 알 수 없는 컨테이너 — 자식 평탄화.
            default:
            {
                if (el.ChildNodes.Length > 0)
                    ProcessChildren(el, target, ctx);
                break;
            }
        }
    }

    private static void ProcessList(IElement el, IList<PdBlock> target, InlineCtx ctx)
    {
        bool isOrdered = el.LocalName == "ol";
        int  start     = 1;
        if (isOrdered && int.TryParse(el.GetAttribute("start"), out var s)) start = s;

        int counter = 0;
        foreach (var child in el.ChildNodes)
        {
            if (child is not IElement li || li.LocalName != "li") continue;
            counter++;

            // GFM 작업 목록 — 첫 자식이 <input type=checkbox> 면 추출.
            bool? checkedState = null;
            var firstInput = li.QuerySelector("input[type=checkbox]");
            if (firstInput is not null)
            {
                checkedState = firstInput.HasAttribute("checked");
                firstInput.Remove();
            }

            int order = li.GetAttribute("value") is { } v && int.TryParse(v, out var ov) ? ov : start + counter - 1;
            var lm = isOrdered
                ? new ListMarker
                {
                    Kind          = ListKind.OrderedDecimal,
                    OrderedNumber = order,
                    Level         = ctx.ListLevel,
                    Checked       = checkedState,
                }
                : new ListMarker
                {
                    Kind    = ListKind.Bullet,
                    Level   = ctx.ListLevel,
                    Checked = checkedState,
                };

            // <li> 내부의 첫 텍스트/인라인은 한 단락으로, 후속 블록(중첩 리스트 등)은 평탄화.
            var firstParagraph = new Paragraph();
            firstParagraph.Style.ListMarker = lm;
            firstParagraph.Style.QuoteLevel = ctx.QuoteLevel;

            bool firstAdded = false;
            foreach (var n in li.ChildNodes)
            {
                if (IsBlockElement(n))
                {
                    if (!firstAdded && firstParagraph.Runs.Count > 0)
                    {
                        target.Add(firstParagraph);
                        firstAdded = true;
                    }
                    ProcessNode(n, target, With(ctx,
                        marker:  lm,
                        listLvl: ctx.ListLevel + 1,
                        quote:   ctx.QuoteLevel));
                }
                else
                {
                    AppendInlineNode(firstParagraph, n, new RunStyle(), url: null);
                }
            }
            if (!firstAdded && firstParagraph.Runs.Count > 0)
                target.Add(firstParagraph);
        }
    }

    private static bool IsBlockElement(INode node) => node is IElement el && el.LocalName switch
    {
        "p" or "div" or "ul" or "ol" or "li" or "blockquote" or "pre" or
        "table" or "thead" or "tbody" or "tfoot" or "tr" or "td" or "th" or
        "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or
        "hr" or "figure" or "section" or "article" or "main" or "aside" or
        "header" or "footer" or "nav" or "details" or "img" => true,
        _ => false,
    };

    private static Paragraph BuildPreCodeParagraph(IElement preEl, InlineCtx ctx)
    {
        var p = new Paragraph();
        p.Style.QuoteLevel   = ctx.QuoteLevel;
        p.Style.ListMarker   = CloneMarker(ctx.Marker);

        // <pre><code class="language-xxx"> 우선 — 언어 추출.
        var inner = preEl.QuerySelector("code");
        var lang  = inner is not null ? ExtractCodeLanguage(inner) : ExtractCodeLanguage(preEl);
        p.Style.CodeLanguage = lang ?? "";

        var text = inner?.TextContent ?? preEl.TextContent;
        // <pre> 의 leading newline 제거 (HTML 관례).
        if (text.StartsWith('\n')) text = text[1..];
        p.AddText(text, MonoStyle());
        return p;
    }

    private static string? ExtractCodeLanguage(IElement el)
    {
        var cls = el.GetAttribute("class") ?? string.Empty;
        foreach (var token in cls.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            if (token.StartsWith("language-", StringComparison.OrdinalIgnoreCase))
                return token[9..];
            if (token.StartsWith("lang-", StringComparison.OrdinalIgnoreCase))
                return token[5..];
        }
        return null;
    }

    private static PdTable BuildTable(IElement tableEl, InlineCtx ctx)
    {
        var t = new PdTable();

        // 표 배경색.
        var tblStyle = tableEl.GetAttribute("style");
        if (StyleProp(tblStyle, "background-color") is { } tblBg
            && TryParseCssColor(tblBg, out var tblBgColor))
            t.BackgroundColor = ColorToHex(tblBgColor);

        // 표 정렬 (margin:auto).
        var marginL = StyleProp(tblStyle, "margin-left");
        var marginR = StyleProp(tblStyle, "margin-right");
        if (marginL == "auto" && marginR == "auto")       t.HAlign = TableHAlign.Center;
        else if (marginL == "auto" && marginR != "auto")  t.HAlign = TableHAlign.Right;

        var rows = tableEl.QuerySelectorAll("tr").ToList();
        int maxCols = rows.Count > 0 ? rows.Max(r => r.QuerySelectorAll("td,th").Count(_ => true)) : 0;
        for (int i = 0; i < maxCols; i++) t.Columns.Add(new TableColumn());

        // <colgroup><col> 에서 열 너비 파싱.
        var colEls = tableEl.QuerySelectorAll("col").ToList();
        for (int i = 0; i < colEls.Count && i < t.Columns.Count; i++)
        {
            var wVal = colEls[i].GetAttribute("width")
                    ?? StyleProp(colEls[i].GetAttribute("style"), "width");
            if (TryParseCssMm(wVal, out var wMm) && wMm > 0)
                t.Columns[i].WidthMm = wMm;
        }

        foreach (var rowEl in rows)
        {
            var row = new PdTableRow();
            // 헤더 행 — 부모가 <thead> 이거나 모든 셀이 <th> 이면 헤더.
            row.IsHeader = rowEl.ParentElement?.LocalName == "thead" ||
                           rowEl.QuerySelectorAll("td,th").All(c => c.LocalName == "th");

            foreach (var cellEl in rowEl.QuerySelectorAll("td,th"))
            {
                var cellStyleStr = cellEl.GetAttribute("style");
                var cell = new PdTableCell
                {
                    TextAlign  = ParseCellAlign(cellEl.GetAttribute("align")
                                              ?? StyleProp(cellStyleStr, "text-align")),
                    ColumnSpan = TryAttrInt(cellEl, "colspan", 1),
                    RowSpan    = TryAttrInt(cellEl, "rowspan", 1),
                };

                // 셀 배경색 (style 속성 또는 bgcolor 속성).
                var bgVal = cellEl.GetAttribute("bgcolor")
                          ?? StyleProp(cellStyleStr, "background-color");
                if (bgVal is not null && TryParseCssColor(bgVal, out var bgColor))
                    cell.BackgroundColor = ColorToHex(bgColor);

                // 셀 안여백 (padding: …mm).
                if (TryParseCssMm(StyleProp(cellStyleStr, "padding"), out var padAll) && padAll > 0)
                {
                    cell.PaddingTopMm    = padAll;
                    cell.PaddingBottomMm = padAll;
                    cell.PaddingLeftMm   = padAll;
                    cell.PaddingRightMm  = padAll;
                }
                if (TryParseCssMm(StyleProp(cellStyleStr, "padding-top"),    out var pt) && pt > 0) cell.PaddingTopMm    = pt;
                if (TryParseCssMm(StyleProp(cellStyleStr, "padding-bottom"), out var pb) && pb > 0) cell.PaddingBottomMm = pb;
                if (TryParseCssMm(StyleProp(cellStyleStr, "padding-left"),   out var pl) && pl > 0) cell.PaddingLeftMm   = pl;
                if (TryParseCssMm(StyleProp(cellStyleStr, "padding-right"),  out var pr) && pr > 0) cell.PaddingRightMm  = pr;

                // 블록 한도는 부모 ctx 와 공유.
                ProcessChildren(cellEl, cell.Blocks, new InlineCtx { Shared = ctx.Shared });
                if (cell.Blocks.Count == 0)
                {
                    var p = new Paragraph();
                    AppendInline(p, cellEl);
                    cell.Blocks.Add(p);
                }
                row.Cells.Add(cell);
            }
            t.Rows.Add(row);
        }

        return t;
    }

    private static ImageBlock BuildImage(IElement imgEl)
    {
        var src   = imgEl.GetAttribute("src") ?? "";
        var alt   = imgEl.GetAttribute("alt");
        var img   = new ImageBlock
        {
            Description  = string.IsNullOrEmpty(alt) ? null : alt,
            ResourcePath = src,
            MediaType    = GuessMediaType(src),
        };

        // data: URI 면 바이너리 디코드 시도.
        if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            int comma = src.IndexOf(',');
            if (comma > 0)
            {
                var meta = src[5..comma];
                var data = src[(comma + 1)..];
                var sep  = meta.IndexOf(';');
                img.MediaType = sep > 0 ? meta[..sep] : meta;
                bool isBase64 = meta.Contains("base64", StringComparison.OrdinalIgnoreCase);
                try
                {
                    img.Data = isBase64
                        ? Convert.FromBase64String(data)
                        : Encoding.UTF8.GetBytes(Uri.UnescapeDataString(data));
                    img.ResourcePath = null; // 바이너리가 있으므로 경로 불필요.
                }
                catch { /* 디코드 실패 시 src 만 보존 */ }
            }
        }

        // 너비/높이 (px → mm 변환은 96 DPI 기준).
        if (TryAttrDouble(imgEl, "width", out var wPx))  img.WidthMm  = wPx * 25.4 / 96.0;
        if (TryAttrDouble(imgEl, "height", out var hPx)) img.HeightMm = hPx * 25.4 / 96.0;

        // CSS float/margin → WrapMode / HAlign.
        var imgStyle = imgEl.GetAttribute("style");
        var floatVal = StyleProp(imgStyle, "float")?.ToLowerInvariant();
        switch (floatVal)
        {
            case "left":  img.WrapMode = ImageWrapMode.WrapRight; break;
            case "right": img.WrapMode = ImageWrapMode.WrapLeft;  break;
        }
        if (floatVal is null)
        {
            var marginL = StyleProp(imgStyle, "margin-left")?.Trim().ToLowerInvariant();
            var marginR = StyleProp(imgStyle, "margin-right")?.Trim().ToLowerInvariant();
            if (marginL == "auto" && marginR == "auto") img.HAlign = ImageHAlign.Center;
            else if (marginL == "auto")                 img.HAlign = ImageHAlign.Right;
        }

        return img;
    }

    private static void BuildFigure(IElement figEl, IList<PdBlock> target, InlineCtx ctx)
    {
        // figcaption → 가까운 img 의 Description 로 흡수, 또는 별도 단락.
        var caption = figEl.QuerySelector("figcaption");
        var imgEl   = figEl.QuerySelector("img");
        if (imgEl is not null)
        {
            var img = BuildImage(imgEl);
            if (caption is not null)
            {
                img.ShowTitle = true;
                img.Title     = caption.TextContent.Trim();
                img.TitlePosition = ImageTitlePosition.Below;
            }
            target.Add(img);

            // figure 의 다른 자식 처리.
            foreach (var n in figEl.ChildNodes)
            {
                if (n == imgEl || n == caption) continue;
                ProcessNode(n, target, ctx);
            }
        }
        else
        {
            // img 없는 figure — 자식만 평탄화.
            ProcessChildren(figEl, target, ctx);
        }
    }

    // ── 인라인 처리 ──────────────────────────────────────────────────────

    private static void AppendInline(Paragraph p, IElement el)
    {
        var initial = ParseInlineStyle(el.GetAttribute("style"));
        ApplyTagStyle(el, ref initial);
        foreach (var n in el.ChildNodes) AppendInlineNode(p, n, initial, url: null);
        if (p.Runs.Count == 0) p.AddText(string.Empty);
    }

    private static void AppendInlineNode(Paragraph p, INode node, RunStyle style, string? url)
    {
        switch (node)
        {
            case IText txt:
                if (txt.Data.Length == 0) return;
                p.Runs.Add(new Run { Text = NormalizeWhitespace(txt.Data), Style = Clone(style), Url = url });
                break;

            case IElement el:
                AppendInlineElement(p, el, style, url);
                break;
        }
    }

    private static void AppendInlineElement(Paragraph p, IElement el, RunStyle parentStyle, string? parentUrl)
    {
        switch (el.LocalName)
        {
            case "br":
                p.Runs.Add(new Run { Text = "\n", Style = Clone(parentStyle), Url = parentUrl });
                return;

            case "img":
                // 인라인 이미지 — 모델 한계로 alt 텍스트 + URL fallback.
                var alt = el.GetAttribute("alt") ?? "";
                var src = el.GetAttribute("src") ?? "";
                p.AddText($"[{alt}]({src})", Clone(parentStyle));
                return;

            case "a":
            {
                var href  = el.GetAttribute("href");
                var style = MergeStyle(parentStyle, el);
                style.Underline = true;
                foreach (var n in el.ChildNodes) AppendInlineNode(p, n, style, href ?? parentUrl);
                return;
            }

            case "sup":
            {
                // Pandoc 스타일 각주/미주 참조: <sup id="fnref-N"> 또는 <sup id="enref-N">
                var supId = el.GetAttribute("id");
                if (supId is not null && supId.StartsWith("fnref-", StringComparison.Ordinal))
                {
                    p.Runs.Add(new Run { FootnoteId = $"fn-{supId[6..]}", Style = Clone(parentStyle) });
                    return;
                }
                if (supId is not null && supId.StartsWith("enref-", StringComparison.Ordinal))
                {
                    p.Runs.Add(new Run { EndnoteId = $"en-{supId[6..]}", Style = Clone(parentStyle) });
                    return;
                }
                // 일반 <sup> — 위첨자 처리.
                var supStyle = MergeStyle(parentStyle, el);
                ApplyTagStyle(el, ref supStyle);
                foreach (var n in el.ChildNodes) AppendInlineNode(p, n, supStyle, parentUrl);
                return;
            }
        }

        var s = MergeStyle(parentStyle, el);
        ApplyTagStyle(el, ref s);
        foreach (var n in el.ChildNodes) AppendInlineNode(p, n, s, parentUrl);
    }

    private static void ApplyTagStyle(IElement el, ref RunStyle s)
    {
        switch (el.LocalName)
        {
            case "strong": case "b":          s.Bold          = true; break;
            case "em": case "i": case "cite": case "var": case "dfn":
                                              s.Italic        = true; break;
            case "u": case "ins":             s.Underline     = true; break;
            case "s": case "strike": case "del":
                                              s.Strikethrough = true; break;
            case "sub":                       s.Subscript     = true; break;
            case "sup":                       s.Superscript   = true; break;
            case "code": case "kbd": case "samp": case "tt":
                                              s.FontFamily    = "Consolas, D2Coding, monospace"; break;
            case "mark":
                s.Background = new Color(0xFF, 0xF5, 0x9D); break;
            case "small":
                if (s.FontSizePt > 1) s.FontSizePt *= 0.85; break;
            case "big":
                s.FontSizePt *= 1.15; break;
            case "abbr":
                // title 속성은 보존 안 함 — 기울임으로 표시.
                s.Italic = true; break;
            case "q":
                // 인용 — 따옴표 마커는 추가하지 않고 스타일만.
                break;
        }
    }

    private static RunStyle MergeStyle(RunStyle parent, IElement el)
    {
        var s = Clone(parent);
        var inline = ParseInlineStyle(el.GetAttribute("style"));
        if (inline.FontFamily is { Length: > 0 } ff) s.FontFamily = ff;
        if (inline.FontSizePt > 0)                   s.FontSizePt = inline.FontSizePt;
        if (inline.Bold)          s.Bold          = true;
        if (inline.Italic)        s.Italic        = true;
        if (inline.Underline)     s.Underline     = true;
        if (inline.Strikethrough) s.Strikethrough = true;
        if (inline.Overline)      s.Overline      = true;
        if (inline.Foreground is { } fg) s.Foreground = fg;
        if (inline.Background is { } bg) s.Background = bg;
        return s;
    }

    private static RunStyle ParseInlineStyle(string? style)
    {
        var s = new RunStyle();
        if (string.IsNullOrWhiteSpace(style)) return s;

        foreach (var declRaw in style.Split(';'))
        {
            var decl = declRaw.Trim();
            int colon = decl.IndexOf(':');
            if (colon <= 0) continue;
            var prop = decl[..colon].Trim().ToLowerInvariant();
            var val  = decl[(colon + 1)..].Trim();

            switch (prop)
            {
                case "font-family":
                    s.FontFamily = val.Trim('"', '\'');
                    break;
                case "font-size":
                    s.FontSizePt = ParseFontSizePt(val);
                    break;
                case "font-weight":
                    s.Bold = val == "bold" || (int.TryParse(val, out var w) && w >= 600);
                    break;
                case "font-style":
                    s.Italic = val == "italic" || val == "oblique";
                    break;
                case "text-decoration":
                case "text-decoration-line":
                    if (val.Contains("underline"))    s.Underline     = true;
                    if (val.Contains("line-through")) s.Strikethrough = true;
                    if (val.Contains("overline"))     s.Overline      = true;
                    break;
                case "color":
                    if (TryParseCssColor(val, out var fg)) s.Foreground = fg;
                    break;
                case "background-color":
                case "background":
                    if (TryParseCssColor(val, out var bg)) s.Background = bg;
                    break;
            }
        }
        return s;
    }

    private static double ParseFontSizePt(string val)
    {
        val = val.Trim().ToLowerInvariant();
        if (val.EndsWith("pt") && double.TryParse(val[..^2], NumberStyles.Any, CultureInfo.InvariantCulture, out var pt)) return pt;
        if (val.EndsWith("px") && double.TryParse(val[..^2], NumberStyles.Any, CultureInfo.InvariantCulture, out var px)) return px * 72.0 / 96.0;
        if (val.EndsWith("em") && double.TryParse(val[..^2], NumberStyles.Any, CultureInfo.InvariantCulture, out var em)) return em * 11.0;
        if (val.EndsWith("rem") && double.TryParse(val[..^3], NumberStyles.Any, CultureInfo.InvariantCulture, out var rem)) return rem * 11.0;
        if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out var num)) return num;
        return 0;
    }

    private static bool TryParseCssColor(string val, out Color color)
    {
        color = default;
        val = val.Trim();
        if (val.StartsWith('#'))
        {
            try { color = Color.FromHex(val); return true; } catch { return false; }
        }
        if (val.StartsWith("rgb", StringComparison.OrdinalIgnoreCase))
        {
            var open  = val.IndexOf('(');
            var close = val.LastIndexOf(')');
            if (open < 0 || close < 0) return false;
            var parts = val[(open + 1)..close].Split(',', StringSplitOptions.TrimEntries);
            if (parts.Length < 3) return false;
            if (!byte.TryParse(parts[0], out var r) ||
                !byte.TryParse(parts[1], out var g) ||
                !byte.TryParse(parts[2], out var b)) return false;
            byte a = 255;
            if (parts.Length >= 4 && double.TryParse(parts[3], NumberStyles.Any, CultureInfo.InvariantCulture, out var an))
                a = (byte)Math.Clamp(an * 255, 0, 255);
            color = new Color(r, g, b, a);
            return true;
        }
        // CSS Level 1/2/3 명명 색상 (주요 141개 중 빈번하게 쓰이는 것).
        Color? named = val.ToLowerInvariant() switch
        {
            // 기본 16색 (CSS Level 1+2)
            "black"    => new Color(  0,   0,   0),
            "silver"   => new Color(192, 192, 192),
            "gray" or "grey"
                       => new Color(128, 128, 128),
            "white"    => new Color(255, 255, 255),
            "maroon"   => new Color(128,   0,   0),
            "red"      => new Color(255,   0,   0),
            "purple"   => new Color(128,   0, 128),
            "fuchsia" or "magenta"
                       => new Color(255,   0, 255),
            "green"    => new Color(  0, 128,   0),
            "lime"     => new Color(  0, 255,   0),
            "olive"    => new Color(128, 128,   0),
            "yellow"   => new Color(255, 255,   0),
            "navy"     => new Color(  0,   0, 128),
            "blue"     => new Color(  0,   0, 255),
            "teal"     => new Color(  0, 128, 128),
            "aqua" or "cyan"
                       => new Color(  0, 255, 255),
            // 추가 CSS3 명명 색상
            "orange"      => new Color(255, 165,   0),
            "orangered"   => new Color(255,  69,   0),
            "gold"        => new Color(255, 215,   0),
            "pink"        => new Color(255, 192, 203),
            "hotpink"     => new Color(255, 105, 180),
            "deeppink"    => new Color(255,  20, 147),
            "crimson"     => new Color(220,  20,  60),
            "darkred"     => new Color(139,   0,   0),
            "tomato"      => new Color(255,  99,  71),
            "coral"       => new Color(255, 127,  80),
            "salmon"      => new Color(250, 128, 114),
            "brown"       => new Color(165,  42,  42),
            "chocolate"   => new Color(210, 105,  30),
            "sienna"      => new Color(160,  82,  45),
            "tan"         => new Color(210, 180, 140),
            "khaki"       => new Color(240, 230, 140),
            "goldenrod"   => new Color(218, 165,  32),
            "darkgoldenrod" => new Color(184, 134,  11),
            "peachpuff"   => new Color(255, 218, 185),
            "bisque"      => new Color(255, 228, 196),
            "wheat"       => new Color(245, 222, 179),
            "beige"       => new Color(245, 245, 220),
            "linen"       => new Color(250, 240, 230),
            "ivory"       => new Color(255, 255, 240),
            "lightyellow" => new Color(255, 255, 224),
            "lightgreen"  => new Color(144, 238, 144),
            "palegreen"   => new Color(152, 251, 152),
            "darkgreen"   => new Color(  0, 100,   0),
            "forestgreen" => new Color( 34, 139,  34),
            "seagreen"    => new Color( 46, 139,  87),
            "mediumseagreen" => new Color( 60, 179, 113),
            "springgreen" => new Color(  0, 255, 127),
            "lawngreen"   => new Color(124, 252,   0),
            "chartreuse"  => new Color(127, 255,   0),
            "yellowgreen" => new Color(154, 205,  50),
            "greenyellow" => new Color(173, 255,  47),
            "turquoise"   => new Color( 64, 224, 208),
            "mediumturquoise" => new Color( 72, 209, 204),
            "darkturquoise" => new Color(  0, 206, 209),
            "lightcyan"   => new Color(224, 255, 255),
            "paleturquoise" => new Color(175, 238, 238),
            "skyblue"     => new Color(135, 206, 235),
            "lightskyblue" => new Color(135, 206, 250),
            "deepskyblue" => new Color(  0, 191, 255),
            "dodgerblue"  => new Color( 30, 144, 255),
            "cornflowerblue" => new Color(100, 149, 237),
            "steelblue"   => new Color( 70, 130, 180),
            "royalblue"   => new Color( 65, 105, 225),
            "darkblue"    => new Color(  0,   0, 139),
            "mediumblue"  => new Color(  0,   0, 205),
            "slateblue"   => new Color(106,  90, 205),
            "mediumpurple" => new Color(147, 112, 219),
            "blueviolet"  => new Color(138,  43, 226),
            "darkviolet"  => new Color(148,   0, 211),
            "darkorchid"  => new Color(153,  50, 204),
            "darkmagenta" => new Color(139,   0, 139),
            "plum"        => new Color(221, 160, 221),
            "violet"      => new Color(238, 130, 238),
            "orchid"      => new Color(218, 112, 214),
            "lavender"    => new Color(230, 230, 250),
            "thistle"     => new Color(216, 191, 216),
            "lightgray" or "lightgrey"
                          => new Color(211, 211, 211),
            "darkgray" or "darkgrey"
                          => new Color(169, 169, 169),
            "dimgray" or "dimgrey"
                          => new Color(105, 105, 105),
            "gainsboro"   => new Color(220, 220, 220),
            "whitesmoke"  => new Color(245, 245, 245),
            "snow"        => new Color(255, 250, 250),
            "ghostwhite"  => new Color(248, 248, 255),
            "aliceblue"   => new Color(240, 248, 255),
            "azure"       => new Color(240, 255, 255),
            "honeydew"    => new Color(240, 255, 240),
            "mintcream"   => new Color(245, 255, 250),
            "seashell"    => new Color(255, 245, 238),
            "floralwhite" => new Color(255, 250, 240),
            "oldlace"     => new Color(253, 245, 230),
            "antiquewhite" => new Color(250, 235, 215),
            "moccasin"    => new Color(255, 228, 181),
            "papayawhip"  => new Color(255, 239, 213),
            "blanchedalmond" => new Color(255, 235, 205),
            "mistyrose"   => new Color(255, 228, 225),
            "lightpink"   => new Color(255, 182, 193),
            "mediumvioletred" => new Color(199,  21, 133),
            "palevioletred" => new Color(219, 112, 147),
            "rosybrown"   => new Color(188, 143, 143),
            "indianred"   => new Color(205,  92,  92),
            "lightcoral"  => new Color(240, 128, 128),
            "lightsalmon" => new Color(255, 160, 122),
            "darksalmon"  => new Color(233, 150, 122),
            "burlywood"   => new Color(222, 184, 135),
            "sandybrown"  => new Color(244, 164,  96),
            "peru"        => new Color(205, 133,  63),
            "saddlebrown" => new Color(139,  69,  19),
            "darkkhaki"   => new Color(189, 183, 107),
            "palegoldenrod" => new Color(238, 232, 170),
            "lemonchiffon" => new Color(255, 250, 205),
            "lightgoldenrodyellow" => new Color(250, 250, 210),
            "cornsilk"    => new Color(255, 248, 220),
            "mediumspringgreen" => new Color(  0, 250, 154),
            "darkseagreen" => new Color(143, 188, 143),
            "lightseagreen" => new Color( 32, 178, 170),
            "darkcyan"    => new Color(  0, 139, 139),
            "cadetblue"   => new Color( 95, 158, 160),
            "powderblue"  => new Color(176, 224, 230),
            "lightblue"   => new Color(173, 216, 230),
            "lightsteelblue" => new Color(176, 196, 222),
            "mediumslateblue" => new Color(123, 104, 238),
            "darkslateblue" => new Color( 72,  61, 139),
            "darkslategray" or "darkslategrey"
                          => new Color( 47,  79,  79),
            "slategray" or "slategrey"
                          => new Color(112, 128, 144),
            "lightslategray" or "lightslategrey"
                          => new Color(119, 136, 153),
            "mediumaquamarine" => new Color(102, 205, 170),
            "aquamarine"  => new Color(127, 255, 212),
            _             => (Color?)null,
        };
        if (named.HasValue) { color = named.Value; return true; }
        return false;
    }

    private static CellTextAlign ParseCellAlign(string? v) => v?.ToLowerInvariant() switch
    {
        "center"  => CellTextAlign.Center,
        "right"   => CellTextAlign.Right,
        "justify" => CellTextAlign.Justify,
        _         => CellTextAlign.Left,
    };

    private static void ApplyBlockAlignment(Paragraph p, IElement el)
    {
        ApplyBlockStyle(p, el);
    }

    /// <summary>블록 요소의 style 속성에서 단락 레이아웃 CSS 를 파싱해 ParagraphStyle 에 반영.</summary>
    private static void ApplyBlockStyle(Paragraph p, IElement el)
    {
        var style = el.GetAttribute("style");
        var align = el.GetAttribute("align") ?? StyleProp(style, "text-align");
        p.Style.Alignment = align?.ToLowerInvariant() switch
        {
            "center"  => Alignment.Center,
            "right"   => Alignment.Right,
            "justify" => Alignment.Justify,
            _         => p.Style.Alignment,
        };

        if (TryParseLineHeight(StyleProp(style, "line-height"), out var lh))
            p.Style.LineHeightFactor = lh;

        if (TryParseCssPt(StyleProp(style, "margin-top"), out var mt))
            p.Style.SpaceBeforePt = mt;
        if (TryParseCssPt(StyleProp(style, "margin-bottom"), out var mb))
            p.Style.SpaceAfterPt = mb;

        if (TryParseCssMm(StyleProp(style, "text-indent"), out var ti))
            p.Style.IndentFirstLineMm = ti;

        // padding-left 우선, 없으면 margin-left (들여쓰기 호환).
        if (TryParseCssMm(StyleProp(style, "padding-left") ?? StyleProp(style, "margin-left"), out var il))
            p.Style.IndentLeftMm = il;
        if (TryParseCssMm(StyleProp(style, "padding-right") ?? StyleProp(style, "margin-right"), out var ir))
            p.Style.IndentRightMm = ir;

        // 강제 페이지 나누기: page-break-before:always (CSS2 legacy) 또는 break-before:page (CSS3).
        var pbv = StyleProp(style, "page-break-before") ?? StyleProp(style, "break-before");
        if (pbv is not null && (pbv.Equals("always", StringComparison.OrdinalIgnoreCase)
                             || pbv.Equals("page",   StringComparison.OrdinalIgnoreCase)))
        {
            p.Style.ForcePageBreakBefore = true;
        }
    }

    private static string? StyleProp(string? style, string prop)
    {
        if (string.IsNullOrEmpty(style)) return null;
        foreach (var decl in style.Split(';'))
        {
            var c = decl.IndexOf(':');
            if (c <= 0) continue;
            if (decl[..c].Trim().Equals(prop, StringComparison.OrdinalIgnoreCase))
                return decl[(c + 1)..].Trim();
        }
        return null;
    }

    // ── 유틸 ────────────────────────────────────────────────────────────

    private static string NormalizeWhitespace(string text)
    {
        // HTML 의 다중 공백/개행은 단일 공백으로 축약 (CSS white-space: normal 모델).
        if (string.IsNullOrEmpty(text)) return text;
        var sb = new StringBuilder(text.Length);
        bool prevSpace = false;
        foreach (var ch in text)
        {
            if (char.IsWhiteSpace(ch))
            {
                if (!prevSpace) { sb.Append(' '); prevSpace = true; }
            }
            else
            {
                sb.Append(ch); prevSpace = false;
            }
        }
        return sb.ToString();
    }

    private static RunStyle MonoStyle() => new() { FontFamily = "Consolas, D2Coding, monospace" };

    // Core 의 정식 RunStyle.Clone() 사용.
    private static RunStyle Clone(RunStyle s) => s.Clone();

    // Core 의 정식 ListMarker.Clone() 사용.
    private static ListMarker? CloneMarker(ListMarker? m) => m?.Clone();

    private static int TryAttrInt(IElement el, string name, int fallback)
        => int.TryParse(el.GetAttribute(name), out var v) ? v : fallback;

    private static bool TryAttrDouble(IElement el, string name, out double v)
        => double.TryParse(el.GetAttribute(name), NumberStyles.Any, CultureInfo.InvariantCulture, out v);

    private static string ColorToHex(Color c)
        => $"#{c.R:X2}{c.G:X2}{c.B:X2}";

    /// <summary>CSS 길이값 → mm 변환. 지원 단위: mm/cm/px/pt/in.</summary>
    private static bool TryParseCssMm(string? val, out double mm)
    {
        mm = 0;
        if (string.IsNullOrWhiteSpace(val)) return false;
        val = val.Trim().ToLowerInvariant();
        if (val.EndsWith("mm")  && double.TryParse(val[..^2],  NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) { mm = v; return true; }
        if (val.EndsWith("cm")  && double.TryParse(val[..^2],  NumberStyles.Any, CultureInfo.InvariantCulture, out v)) { mm = v * 10; return true; }
        if (val.EndsWith("px")  && double.TryParse(val[..^2],  NumberStyles.Any, CultureInfo.InvariantCulture, out v)) { mm = v * 25.4 / 96.0; return true; }
        if (val.EndsWith("pt")  && double.TryParse(val[..^2],  NumberStyles.Any, CultureInfo.InvariantCulture, out v)) { mm = v * 25.4 / 72.0; return true; }
        if (val.EndsWith("in")  && double.TryParse(val[..^2],  NumberStyles.Any, CultureInfo.InvariantCulture, out v)) { mm = v * 25.4; return true; }
        return false;
    }

    /// <summary>CSS 길이값 → pt 변환. 지원 단위: pt/px/mm/cm/in.</summary>
    private static bool TryParseCssPt(string? val, out double pt)
    {
        pt = 0;
        if (string.IsNullOrWhiteSpace(val)) return false;
        val = val.Trim().ToLowerInvariant();
        if (val.EndsWith("pt") && double.TryParse(val[..^2], NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) { pt = v; return true; }
        if (val.EndsWith("px") && double.TryParse(val[..^2], NumberStyles.Any, CultureInfo.InvariantCulture, out v)) { pt = v * 72.0 / 96.0; return true; }
        if (val.EndsWith("mm") && double.TryParse(val[..^2], NumberStyles.Any, CultureInfo.InvariantCulture, out v)) { pt = v * 72.0 / 25.4; return true; }
        if (val.EndsWith("cm") && double.TryParse(val[..^2], NumberStyles.Any, CultureInfo.InvariantCulture, out v)) { pt = v * 720.0 / 25.4; return true; }
        if (val.EndsWith("in") && double.TryParse(val[..^2], NumberStyles.Any, CultureInfo.InvariantCulture, out v)) { pt = v * 72.0; return true; }
        return false;
    }

    /// <summary>CSS line-height → 배율 변환. 단위 없음(배율)/em/% 지원.</summary>
    private static bool TryParseLineHeight(string? val, out double factor)
    {
        factor = 0;
        if (string.IsNullOrWhiteSpace(val)) return false;
        val = val.Trim().ToLowerInvariant();
        if (val == "normal") { factor = 1.2; return true; }
        // 단위 없음 → 직접 배율.
        if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) && v > 0) { factor = v; return true; }
        // em = 배율과 동일 의미.
        if (val.EndsWith("em") && double.TryParse(val[..^2], NumberStyles.Any, CultureInfo.InvariantCulture, out v) && v > 0) { factor = v; return true; }
        // % → 배율.
        if (val.EndsWith('%')  && double.TryParse(val[..^1], NumberStyles.Any, CultureInfo.InvariantCulture, out v) && v > 0) { factor = v / 100.0; return true; }
        return false;
    }

    private static string GuessMediaType(string? url)
    {
        if (string.IsNullOrEmpty(url)) return "application/octet-stream";
        var ext = Path.GetExtension(url).ToLowerInvariant();
        return ext switch
        {
            ".png"  => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif"  => "image/gif",
            ".bmp"  => "image/bmp",
            ".tif" or ".tiff" => "image/tiff",
            ".webp" => "image/webp",
            ".svg"  => "image/svg+xml",
            _       => "application/octet-stream",
        };
    }
}
