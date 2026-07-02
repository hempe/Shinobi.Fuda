<p align="center">
  <img src="icon.png" width="96" height="96" alt="Shinobi.Fuda Logo">
</p>

# Shinobi.Fuda

Docx placeholder-replacement engine: `[[Scalar.Path]]` anywhere, `[[Collection.Property]]`
inside a table row (row clones per item), `[[Repeat:Name]]...[[EndRepeat:Name]]` for a
repeatable multi-element region (row clones aren't enough when the repeating unit spans
several paragraphs/tables — e.g. one templated section per customer card).

```
src/Shinobi.Fuda/        the actual library
samples/Shinobi.Fuda.Samples/   runnable console app + sample invoice template
tests/Shinobi.Fuda.Tests/       xunit tests, self-contained (no external files)
```

## Build & test

```
dotnet restore
dotnet build
dotnet test
```

## Try it end to end

```
cd samples/Shinobi.Fuda.Samples
dotnet run
```

Renders `templates/invoice-template.docx` (generated from the reference screenshot,
two fuel cards) against `SampleInvoiceData.Build()` and writes `generated-invoice.docx`
next to the built exe. Open it in Word to check the block-repeat output looks right —
particularly table borders/shading on the second card, since `CloneNode(deep: true)`
on a `Table` is the part most worth double-checking against Word's own rendering quirks.

## Wiring in

`src/Shinobi.Fuda` has no dependency on anything project-specific — the only package
reference is `DocumentFormat.OpenXml`. Reference it from whichever project owns
invoice generation, and replace the example `InvoiceDto` in `samples/` with your real
invoice model — `TemplateData.FromDto` works on any DTO shape, the reflection rules
are documented in `src/Shinobi.Fuda/TemplateData.cs`.
