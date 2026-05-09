using PolyDonky.Core;

namespace PolyDonky.Core.Tests;

public class HeaderFooterTokensTests
{
    private static readonly DateTime FixedNow =
        new(2026, 5, 5, 14, 30, 0, DateTimeKind.Local);

    private static HeaderFooterTokens.Context Ctx(int page = 3, int total = 10) => new()
    {
        PageNumber = page,
        TotalPages = total,
        Now        = FixedNow,
        Title      = "라운드트립 보고서",
        Author     = "노진문",
        FileName   = "report",
    };

    [Theory]
    [InlineData("{PAGE}", "3")]
    [InlineData("{페이지}", "3")]
    [InlineData("{NUMPAGES}", "10")]
    [InlineData("{전체페이지}", "10")]
    [InlineData("{DATE}", "2026-05-05")]
    [InlineData("{날짜}", "2026-05-05")]
    [InlineData("{TIME}", "14:30")]
    [InlineData("{시간}", "14:30")]
    [InlineData("{TITLE}", "라운드트립 보고서")]
    [InlineData("{제목}", "라운드트립 보고서")]
    [InlineData("{AUTHOR}", "노진문")]
    [InlineData("{저자}", "노진문")]
    [InlineData("{FILENAME}", "report")]
    [InlineData("{파일명}", "report")]
    public void Resolve_SingleToken(string template, string expected)
    {
        Assert.Equal(expected, HeaderFooterTokens.Resolve(template, Ctx()));
    }

    [Fact]
    public void Resolve_MixedTextAndTokens()
    {
        Assert.Equal(
            "- 3 / 10 -",
            HeaderFooterTokens.Resolve("- {PAGE} / {NUMPAGES} -", Ctx()));
    }

    [Fact]
    public void Resolve_Korean_PageOfTotal()
    {
        Assert.Equal(
            "3 / 10 페이지",
            HeaderFooterTokens.Resolve("{페이지} / {전체페이지} 페이지", Ctx()));
    }

    [Fact]
    public void Resolve_TokenNamesAreCaseInsensitive()
    {
        Assert.Equal("3", HeaderFooterTokens.Resolve("{Page}", Ctx()));
        Assert.Equal("3", HeaderFooterTokens.Resolve("{page}", Ctx()));
        Assert.Equal("10", HeaderFooterTokens.Resolve("{NumPages}", Ctx()));
    }

    [Fact]
    public void Resolve_UnknownToken_LeftAsLiteral()
    {
        // 알 수 없는 토큰은 보존 — 사용자 의도(미래 토큰 또는 단순 중괄호 텍스트) 보호.
        Assert.Equal(
            "값: {UNKNOWN_TOKEN} / 3",
            HeaderFooterTokens.Resolve("값: {UNKNOWN_TOKEN} / {PAGE}", Ctx()));
    }

    [Fact]
    public void Resolve_EmptyOrNull_ReturnsEmpty()
    {
        Assert.Equal(string.Empty, HeaderFooterTokens.Resolve(null, Ctx()));
        Assert.Equal(string.Empty, HeaderFooterTokens.Resolve("", Ctx()));
    }

    [Fact]
    public void Resolve_EscapedBraces_ProduceLiteralBraces()
    {
        Assert.Equal("{PAGE}", HeaderFooterTokens.Resolve(@"\{PAGE\}", Ctx()));
        Assert.Equal("{ 3 }",  HeaderFooterTokens.Resolve(@"\{ {PAGE} \}", Ctx()));
    }

    [Fact]
    public void Resolve_UnterminatedBrace_LeftAsLiteral()
    {
        Assert.Equal(
            "{PAGE 시작만",
            HeaderFooterTokens.Resolve("{PAGE 시작만", Ctx()));
    }

    [Fact]
    public void Resolve_MissingMetadata_ProducesEmptyString()
    {
        var ctx = new HeaderFooterTokens.Context
        {
            PageNumber = 1,
            TotalPages = 1,
            Now = FixedNow,
            // Title/Author/FileName 모두 null
        };
        Assert.Equal(
            "[][]",
            HeaderFooterTokens.Resolve("[{TITLE}][{AUTHOR}]", ctx));
    }

    [Fact]
    public void Resolve_TokenWithSurroundingWhitespace_StillResolved()
    {
        // {  PAGE  } 같이 사용자가 공백을 둔 경우도 인식한다 — 편집 편의.
        Assert.Equal("3", HeaderFooterTokens.Resolve("{  PAGE  }", Ctx()));
    }
}
