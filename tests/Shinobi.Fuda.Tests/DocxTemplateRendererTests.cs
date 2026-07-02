using Xunit;

namespace Shinobi.Fuda.Tests;

public class DocxTemplateRendererTests
{
    [Fact]
    public void ReplacesScalarPlaceholders()
    {
        var template = TestDocx.Create(
            TestDocx.Para("Hello [[Customer.Name]], invoice [[Invoice.Number]]"));

        var data = new TemplateData();
        data.Scalars["Customer.Name"] = "John Smith";
        data.Scalars["Invoice.Number"] = "INV-1";

        var text = TestDocx.GetAllText(new DocxTemplateRenderer().Render(template, data));

        Assert.Contains("Hello John Smith, invoice INV-1", text);
    }

    [Fact]
    public void ClonesTableRowPerCollectionItem()
    {
        var templateRow = TestDocx.Row("[[Records.Name]]", "[[Records.Qty]]");
        var table = TestDocx.SimpleTable(templateRow);
        var template = TestDocx.Create(table);

        var data = new TemplateData();
        data.Collections["Records"] = new List<Dictionary<string, string>>
        {
            new(StringComparer.OrdinalIgnoreCase) { ["Name"] = "Apple", ["Qty"] = "5" },
            new(StringComparer.OrdinalIgnoreCase) { ["Name"] = "Orange", ["Qty"] = "3" },
        };

        var text = TestDocx.GetAllText(new DocxTemplateRenderer().Render(template, data));

        Assert.Contains("Apple", text);
        Assert.Contains("Orange", text);
        Assert.DoesNotContain("[[", text);
    }

    [Fact]
    public void EmptyCollectionRemovesTemplateRow()
    {
        var templateRow = TestDocx.Row("[[Records.Name]]", "[[Records.Qty]]");
        var table = TestDocx.SimpleTable(templateRow);
        var template = TestDocx.Create(table);

        var data = new TemplateData();
        data.Collections["Records"] = new List<Dictionary<string, string>>();

        var text = TestDocx.GetAllText(new DocxTemplateRenderer().Render(template, data));

        Assert.DoesNotContain("[[", text);
        Assert.Equal(string.Empty, text);
    }

    [Fact]
    public void RepeatsBlockAndResolvesNestedCollectionPerItem()
    {
        var startMarker = TestDocx.Para("[[Repeat:Cards as Card]]");
        var header = TestDocx.Para("Card [[Card.Number]] total [[Card.Total]]");
        var txTable = TestDocx.SimpleTable(TestDocx.Row("[[Card.Transactions.Product]]", "[[Card.Transactions.Amount]]"));
        var endMarker = TestDocx.Para("[[EndRepeat:Cards]]");

        var template = TestDocx.Create(startMarker, header, txTable, endMarker);

        var card1 = new TemplateData();
        card1.Scalars["Number"] = "1124";
        card1.Scalars["Total"] = "77.68";
        card1.Collections["Transactions"] = new List<Dictionary<string, string>>
        {
            new(StringComparer.OrdinalIgnoreCase) { ["Product"] = "Bleifrei 95", ["Amount"] = "77.68" },
        };

        var card2 = new TemplateData();
        card2.Scalars["Number"] = "1130";
        card2.Scalars["Total"] = "151.64";
        card2.Collections["Transactions"] = new List<Dictionary<string, string>>
        {
            new(StringComparer.OrdinalIgnoreCase) { ["Product"] = "Diesel", ["Amount"] = "70.61" },
            new(StringComparer.OrdinalIgnoreCase) { ["Product"] = "Diesel", ["Amount"] = "81.03" },
        };

        var data = new TemplateData();
        data.Blocks["Cards"] = new List<TemplateData> { card1, card2 };

        var text = TestDocx.GetAllText(new DocxTemplateRenderer().Render(template, data));

        Assert.Contains("Card 1124 total 77.68", text);
        Assert.Contains("Card 1130 total 151.64", text);
        Assert.Contains("Bleifrei 95", text);
        Assert.Contains("Diesel", text);
        Assert.DoesNotContain("[[", text);
        // No leftover marker/template-row text either.
        Assert.DoesNotContain("Repeat", text);
    }

    [Fact]
    public void UnknownScalarThrowsValidationException()
    {
        var template = TestDocx.Create(TestDocx.Para("Hello [[Unknown.Thing]]"));
        var data = new TemplateData();

        var ex = Assert.Throws<TemplateValidationException>(
            () => new DocxTemplateRenderer().Render(template, data));

        Assert.Contains(ex.Errors, e => e.Contains("Unknown.Thing"));
    }

    [Fact]
    public void MixedCollectionsInSameRowIsRejected()
    {
        var row = TestDocx.Row("[[Records.Name]]", "[[Totals.Amount]]");
        var table = TestDocx.SimpleTable(row);
        var template = TestDocx.Create(table);

        var data = new TemplateData();
        data.Collections["Records"] = new List<Dictionary<string, string>>
        {
            new(StringComparer.OrdinalIgnoreCase) { ["Name"] = "x" },
        };
        data.Collections["Totals"] = new List<Dictionary<string, string>>
        {
            new(StringComparer.OrdinalIgnoreCase) { ["Amount"] = "1" },
        };

        var ex = Assert.Throws<TemplateValidationException>(
            () => new DocxTemplateRenderer().Render(template, data));

        Assert.Contains(ex.Errors, e => e.Contains("only one collection"));
    }

    [Fact]
    public void PageBreakBeforeInsideRepeatBlockIsStrippedAndWarned()
    {
        var startMarker = TestDocx.Para("[[Repeat:Cards as Card]]");
        var header = TestDocx.ParaWithPageBreakBefore("Card [[Card.Number]]");
        var endMarker = TestDocx.Para("[[EndRepeat:Cards]]");
        var template = TestDocx.Create(startMarker, header, endMarker);

        var card1 = new TemplateData();
        card1.Scalars["Number"] = "1";
        var card2 = new TemplateData();
        card2.Scalars["Number"] = "2";

        var data = new TemplateData();
        data.Blocks["Cards"] = new List<TemplateData> { card1, card2 };

        var generated = new DocxTemplateRenderer().Render(template, data, out var warnings);

        Assert.False(TestDocx.HasAnyPageBreak(generated));
        Assert.Contains(warnings, w => w.Contains("page break before", StringComparison.OrdinalIgnoreCase)
                                     && w.Contains("Cards"));
        Assert.Contains("Card 1", TestDocx.GetAllText(generated));
        Assert.Contains("Card 2", TestDocx.GetAllText(generated));
    }

    [Fact]
    public void ManualPageBreakInsideRepeatBlockIsStrippedAndWarned()
    {
        var startMarker = TestDocx.Para("[[Repeat:Cards as Card]]");
        var header = TestDocx.ParaWithManualPageBreak("Card [[Card.Number]]");
        var endMarker = TestDocx.Para("[[EndRepeat:Cards]]");
        var template = TestDocx.Create(startMarker, header, endMarker);

        var data = new TemplateData();
        data.Blocks["Cards"] = new List<TemplateData>
        {
            new() { Scalars = { ["Number"] = "1" } },
            new() { Scalars = { ["Number"] = "2" } },
        };

        var generated = new DocxTemplateRenderer().Render(template, data, out var warnings);

        Assert.False(TestDocx.HasAnyPageBreak(generated));
        Assert.Contains(warnings, w => w.Contains("manual page break", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void PageBreakBeforeInsideCollectionRowIsStrippedAndWarned()
    {
        var templateRow = TestDocx.Row("[[Records.Name]]");
        templateRow.Elements<TableCell>().First().Append(TestDocx.ParaWithPageBreakBefore(""));
        var table = TestDocx.SimpleTable(templateRow);
        var template = TestDocx.Create(table);

        var data = new TemplateData();
        data.Collections["Records"] = new List<Dictionary<string, string>>
        {
            new(StringComparer.OrdinalIgnoreCase) { ["Name"] = "Apple" },
            new(StringComparer.OrdinalIgnoreCase) { ["Name"] = "Orange" },
        };

        var generated = new DocxTemplateRenderer().Render(template, data, out var warnings);

        Assert.False(TestDocx.HasAnyPageBreak(generated));
        Assert.Contains(warnings, w => w.Contains("page break before", StringComparison.OrdinalIgnoreCase)
                                     && w.Contains("Records"));
    }

    [Fact]
    public void NoWarningsWhenTemplateHasNoPageBreaks()
    {
        var template = TestDocx.Create(TestDocx.Para("Hello [[Customer.Name]]"));
        var data = new TemplateData();
        data.Scalars["Customer.Name"] = "Jane";

        new DocxTemplateRenderer().Render(template, data, out var warnings);

        Assert.Empty(warnings);
    }

    [Fact]
    public void UnmatchedRepeatMarkerIsReportedNotSilentlyIgnored()
    {
        var template = TestDocx.Create(
            TestDocx.Para("[[Repeat:Cards as Card]]"),
            TestDocx.Para("Card [[Card.Number]]"));
        // no [[EndRepeat:Cards]]

        var data = new TemplateData();
        data.Blocks["Cards"] = new List<TemplateData>();

        var ex = Assert.Throws<TemplateValidationException>(
            () => new DocxTemplateRenderer().Render(template, data));

        Assert.Contains(ex.Errors, e => e.Contains("no matching [[EndRepeat:Cards]]"));
    }
}
