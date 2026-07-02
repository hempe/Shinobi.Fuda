using Shinobi.Fuda.Examples;

namespace Shinobi.Fuda.Samples;

/// <summary>
/// Builds an InvoiceDto with the same numbers as the reference screenshot
/// (two fuel cards, one per card total, plus a cross-card invoice summary).
/// </summary>
public static class SampleInvoiceData
{
    public static InvoiceDto Build()
    {
        var card1 = new CardDto
        {
            Number = "1124",
            HolderName = "Card Holder 1",
            Total = 77.68m,
            Transactions = new List<TransactionDto>
            {
                new()
                {
                    Date = new DateTime(2024, 10, 18),
                    Time = new DateTime(2024, 10, 18, 8, 53, 0),
                    Number = "1023",
                    Product = "Bleifrei 95",
                    TaxCode = "C",
                    Quantity = 47.66m,
                    Price = 1.670m,
                    Amount = 77.68m,
                },
            },
            Summary = new List<ProductSummaryDto>
            {
                new() { Product = "Bleifrei 95", Quantity = 47.66m, Amount = 79.59m, Discount = 1.91m },
            },
            Vat = new List<VatLineDto>
            {
                new() { Rate = 0.081m, Gross = 77.68m, Tax = 6.29m, Net = 71.39m },
            },
        };

        var card2 = new CardDto
        {
            Number = "1130",
            HolderName = "Card Holder 2",
            Total = 151.64m,
            Transactions = new List<TransactionDto>
            {
                new()
                {
                    Date = new DateTime(2024, 10, 9),
                    Time = new DateTime(2024, 10, 9, 12, 7, 0),
                    Number = "1035",
                    Product = "Diesel",
                    TaxCode = "C",
                    Quantity = 41.05m,
                    Price = 1.760m,
                    Amount = 70.61m,
                },
                new()
                {
                    Date = new DateTime(2024, 10, 21),
                    Time = new DateTime(2024, 10, 21, 14, 40, 0),
                    Number = "1035",
                    Product = "Diesel",
                    TaxCode = "C",
                    Quantity = 47.11m,
                    Price = 1.760m,
                    Amount = 81.03m,
                },
            },
            Summary = new List<ProductSummaryDto>
            {
                new() { Product = "Diesel", Quantity = 88.16m, Amount = 155.16m, Discount = 3.52m },
            },
            Vat = new List<VatLineDto>
            {
                new() { Rate = 0.081m, Gross = 151.64m, Tax = 12.28m, Net = 139.36m },
            },
        };

        return new InvoiceDto
        {
            CustomerName = "Muster Tankstelle AG",
            InvoiceNumber = "INV-2024-10-0042",
            InvoiceDate = new DateTime(2024, 10, 31),
            Cards = new List<CardDto> { card1, card2 },
            InvoiceSummary = new List<ProductSummaryDto>
            {
                new() { Product = "Bleifrei 95", Quantity = 47.66m, Amount = 79.59m, Discount = 1.91m },
                new() { Product = "Diesel", Quantity = 88.16m, Amount = 155.16m, Discount = 3.52m },
                new() { Product = "Administrativkosten CHF 3.5", Quantity = 1m, Amount = 3.5m, Discount = 0m },
                new() { Product = "Rundung", Quantity = 0m, Amount = 0m, Discount = 0m },
            },
            InvoiceVat = new List<VatLineDto>
            {
                new() { Rate = 0.081m, Gross = 77.68m, Tax = 6.29m, Net = 71.39m },
                new() { Rate = 0.081m, Gross = 151.64m, Tax = 12.28m, Net = 139.36m },
                new() { Rate = 0.081m, Gross = 3.5m, Tax = 0.26m, Net = 3.24m },
                new() { Rate = 0m, Gross = 0m, Tax = 0m, Net = 0m },
            },
        };
    }
}
