using System.Globalization;

namespace Shinobi.Fuda.Examples;

// Template markers used for this DTO shape:
//
//   [[Repeat:Cards as Card]]
//   Kundenkarte [[Card.Number]] – [[Card.HolderName]]                [[Card.Total]] CHF
//
//   Datum       Zeit    Nummer   Produkt   Steuer   Menge   Preis   Betrag
//   [[Card.Transactions.Date]]  [[Card.Transactions.Time]]  [[Card.Transactions.Number]]
//   [[Card.Transactions.Product]]  [[Card.Transactions.TaxCode]]  [[Card.Transactions.Quantity]]
//   [[Card.Transactions.Price]]  [[Card.Transactions.Amount]]
//
//   Übersicht                          MwSt.-Rekapitulation
//   Produkt  Menge  Betrag  Rabatt     Satz  Brutto  Steuer  Netto
//   [[Card.Summary.Product]] [[Card.Summary.Quantity]] [[Card.Summary.Amount]] [[Card.Summary.Discount]]
//   [[Card.Vat.Rate]] [[Card.Vat.Gross]] [[Card.Vat.Tax]] [[Card.Vat.Net]]
//   [[EndRepeat:Cards]]
//
// Only ONE table row per table is the "template row" — the renderer clones it once
// per Transactions/Summary/Vat item, same as any other collection-in-a-table.

public sealed class TransactionDto
{
    [TemplateFormat("dd.MM.yyyy")]
    public DateTime Date { get; set; }

    [TemplateFormat("HH:mm")]
    public DateTime Time { get; set; }

    public string Number { get; set; } = "";
    public string Product { get; set; } = "";
    public string TaxCode { get; set; } = "";

    [TemplateFormat("N2")]
    public decimal Quantity { get; set; }

    [TemplateFormat("N3")]
    public decimal Price { get; set; }

    [TemplateFormat("N2")]
    public decimal Amount { get; set; }
}

public sealed class ProductSummaryDto
{
    public string Product { get; set; } = "";

    [TemplateFormat("N2")]
    public decimal Quantity { get; set; }

    [TemplateFormat("N2")]
    public decimal Amount { get; set; }

    [TemplateFormat("N2")]
    public decimal Discount { get; set; }
}

public sealed class VatLineDto
{
    [TemplateFormat("P2")]
    public decimal Rate { get; set; }

    [TemplateFormat("N2")]
    public decimal Gross { get; set; }

    [TemplateFormat("N2")]
    public decimal Tax { get; set; }

    [TemplateFormat("N2")]
    public decimal Net { get; set; }
}

/// <summary>
/// One customer card ("Kundenkarte"). This is what makes "Cards" a Block rather than
/// a flat Collection: it has its own nested lists (Transactions, Summary, Vat).
/// </summary>
public sealed class CardDto
{
    public string Number { get; set; } = "";
    public string HolderName { get; set; } = "";

    [TemplateFormat("N2")]
    public decimal Total { get; set; }

    public List<TransactionDto> Transactions { get; set; } = new();
    public List<ProductSummaryDto> Summary { get; set; } = new();
    public List<VatLineDto> Vat { get; set; } = new();
}

public sealed class InvoiceDto
{
    public string CustomerName { get; set; } = "";
    public string InvoiceNumber { get; set; } = "";

    [TemplateFormat("dd.MM.yyyy")]
    public DateTime InvoiceDate { get; set; }

    public List<CardDto> Cards { get; set; } = new();

    // Cross-card aggregate ("Rechnungsübersicht" at the bottom of your layout).
    // This is NOT computed by the renderer — sum it in your service before building
    // the DTO, same as any other business logic. Kept as a flat Collection/Records-style
    // table (Records/Totals pattern), not a Block, since it has no further nesting.
    public List<ProductSummaryDto> InvoiceSummary { get; set; } = new();
    public List<VatLineDto> InvoiceVat { get; set; } = new();
}

public sealed class InvoiceRenderService
{
    private readonly DocxTemplateRenderer _renderer = new();

    public byte[] RenderInvoice(byte[] templateBytes, InvoiceDto invoice, CultureInfo? culture = null)
    {
        var data = TemplateData.FromDto(invoice, culture ?? CultureInfo.GetCultureInfo("de-CH"));
        return _renderer.Render(templateBytes, data);
    }
}

/*
try
{
    var generated = service.RenderInvoice(templateBytes, invoice);
    await File.WriteAllBytesAsync("invoice.docx", generated);
}
catch (TemplateValidationException ex)
{
    foreach (var e in ex.Errors) logger.LogWarning(e);
}
*/
