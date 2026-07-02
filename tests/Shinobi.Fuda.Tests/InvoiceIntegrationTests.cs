using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using Shinobi.Fuda.Examples;
using Shinobi.Fuda.Samples;
using Xunit;

namespace Shinobi.Fuda.Tests;

/// <summary>
/// End-to-end: the real sample template against the real (complex, nested) InvoiceDto,
/// as opposed to the synthetic docx fixtures used by the other renderer tests. Catches
/// mismatches between template placeholders and the DTO's flattened paths that unit
/// tests built on throwaway documents never would.
/// </summary>
public class InvoiceIntegrationTests
{
    private static string TemplatePath =>
        Path.Combine(AppContext.BaseDirectory, "templates", "invoice-template.docx");

    [Fact]
    public void RendersRealTemplateAgainstComplexInvoiceDto()
    {
        var templateBytes = File.ReadAllBytes(TemplatePath);
        var invoice = SampleInvoiceData.Build();
        var service = new InvoiceRenderService();

        var generated = service.RenderInvoice(templateBytes, invoice);
        var text = GetAllText(generated);

        Assert.Contains("Muster Tankstelle AG", text);
        Assert.Contains("INV-2024-10-0042", text);
        Assert.Contains("31.10.2024", text);

        // Both cards rendered with their own data, not just the first one repeated.
        Assert.Contains("Card Holder 1", text);
        Assert.Contains("Card Holder 2", text);
        Assert.Contains("1124", text);
        Assert.Contains("1130", text);

        Assert.Contains("77.68", text);
        Assert.Contains("70.61", text);
        Assert.Contains("81.03", text);

        // "Diesel" appears in card 2's two transactions, its own summary row, and the
        // cross-card invoice summary -- confirms rows clone per item, not once per block.
        Assert.Equal(4, Regex.Matches(text, "Diesel").Count);
        Assert.Equal(3, Regex.Matches(text, "Bleifrei 95").Count);

        Assert.DoesNotContain("[[", text);
        Assert.DoesNotContain("Repeat", text);
    }

    private static string GetAllText(byte[] docBytes)
    {
        using var stream = new MemoryStream(docBytes);
        using var doc = WordprocessingDocument.Open(stream, isEditable: false);
        return doc.MainDocumentPart!.Document.Body!.InnerText;
    }
}
