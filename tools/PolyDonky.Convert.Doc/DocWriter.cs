using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using PolyDonky.Core;

namespace PolyDonky.Convert.Doc;

/// <summary>
/// IWPF → DOC (97-2003) 변환기. 현재 단계: 텍스트만.
/// </summary>
public class DocWriter
{
    public void Write(PolyDonkyument doc, Stream output)
    {
        var hwpfDoc = new HWPFDocument();
        var range = hwpfDoc.GetRange();

        foreach (var section in doc.Sections)
        {
            foreach (var block in section.Blocks)
            {
                if (block is Paragraph para)
                {
                    WriteBlock(range, para);
                }
                // TODO: Table, ImageBlock, ShapeObject, TextBoxObject 등은 나중에
            }
        }

        hwpfDoc.Write(output);
    }

    private void WriteBlock(Range range, Paragraph para)
    {
        // 문단 생성
        var hwpfPara = range.InsertBefore(HWPFDocument.CreateParagraph());

        // Run들의 텍스트 추출해서 추가
        foreach (var run in para.Runs)
        {
            if (!string.IsNullOrEmpty(run.Text))
            {
                var hwpfRun = hwpfPara.InsertBefore(HWPFDocument.CreateRun(run.Text));

                // TODO: 텍스트 스타일 (굵게, 기울임 등) 적용
                // run.Style.IsBold, IsItalic 등 확인 후 hwpfRun에 적용
            }
        }
    }
}
