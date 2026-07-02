using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Shinobi.Fuda.Tests;

/// <summary>
/// Builds throwaway .docx byte[] fixtures directly via Open XML, so tests don't
/// depend on any external template file.
/// </summary>
internal static class TestDocx
{
    public static byte[] Create(params OpenXmlElement[] bodyContent)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = new Body();
            foreach (var element in bodyContent)
                body.Append(element);
            mainPart.Document.Append(body);
            mainPart.Document.Save();
        }

        return stream.ToArray();
    }

    public static Paragraph Para(string text) =>
        new(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    public static Paragraph ParaWithPageBreakBefore(string text) =>
        new(
            new ParagraphProperties(new PageBreakBefore()),
            new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    public static Paragraph ParaWithManualPageBreak(string text) =>
        new(new Run(new Break { Type = BreakValues.Page }, new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    public static bool HasAnyPageBreak(byte[] docBytes)
    {
        using var stream = new MemoryStream(docBytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        var body = doc.MainDocumentPart!.Document.Body!;

        return body.Descendants<PageBreakBefore>().Any(b => b.Val is null || b.Val.Value)
            || body.Descendants<Break>().Any(b => b.Type?.Value == BreakValues.Page);
    }

    public static Table SimpleTable(params TableRow[] rows)
    {
        var table = new Table();
        foreach (var row in rows)
            table.Append(row);
        return table;
    }

    public static TableRow Row(params string[] cellTexts)
    {
        var row = new TableRow();
        foreach (var text in cellTexts)
            row.Append(new TableCell(Para(text)));
        return row;
    }

    public static string GetAllText(byte[] docBytes)
    {
        using var stream = new MemoryStream(docBytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        return doc.MainDocumentPart!.Document.Body!.InnerText;
    }
}
