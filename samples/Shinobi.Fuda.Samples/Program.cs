using Shinobi.Fuda;
using Shinobi.Fuda.Examples;
using Shinobi.Fuda.Samples;

var templatePath = args.Length > 0 ? args[0] : Path.Combine(AppContext.BaseDirectory, "templates", "invoice-template.docx");
var outputPath = args.Length > 1 ? args[1] : Path.Combine(AppContext.BaseDirectory, "generated-invoice.docx");

if (!File.Exists(templatePath))
{
    Console.Error.WriteLine($"Template not found: {templatePath}");
    return 1;
}

var templateBytes = await File.ReadAllBytesAsync(templatePath);
var invoice = SampleInvoiceData.Build();
var service = new InvoiceRenderService();

try
{
    var generated = service.RenderInvoice(templateBytes, invoice);
    await File.WriteAllBytesAsync(outputPath, generated);
    Console.WriteLine($"Rendered invoice written to: {outputPath}");
    return 0;
}
catch (TemplateValidationException ex)
{
    Console.Error.WriteLine("Template validation failed:");
    foreach (var error in ex.Errors)
        Console.Error.WriteLine($"  - {error}");
    return 1;
}
