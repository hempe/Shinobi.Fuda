using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Shinobi.Fuda;

/// <summary>
/// Thrown when a template references unknown placeholders/collections/blocks, or
/// violates a structural rule (one collection per row, matched Repeat/EndRepeat).
/// Carries every problem found in a single render pass.
/// </summary>
public sealed class TemplateValidationException : Exception
{
    public IReadOnlyList<string> Errors { get; }

    public TemplateValidationException(IReadOnlyList<string> errors)
        : base("Template validation failed:" + Environment.NewLine + string.Join(Environment.NewLine, errors))
    {
        Errors = errors;
    }
}

/// <summary>
/// Renders a .docx template against a <see cref="TemplateData"/> instance.
///
///  - [[Scalar.Path]] may appear anywhere: body paragraphs, headers, footers, table cells.
///  - [[Collection.Property]] is only valid inside a table row; that row is cloned once
///    per item in the collection and the original template row is removed.
///  - [[Repeat:Name]] ... [[EndRepeat:Name]] (each marker alone on its own paragraph)
///    delimits an arbitrary range of sibling content — paragraphs, tables, whatever —
///    that gets cloned once per item in the "Name" block. Placeholders inside the range
///    are namespaced under an alias: [[Repeat:Cards as Card]] means placeholders in that
///    range are written [[Card.Number]], [[Card.Transactions.Date]], etc. If "as Alias"
///    is omitted, the alias defaults to the block name itself.
///  - A table row may reference exactly one collection.
///  - Unknown placeholders/collections/blocks and mixed-collection rows are reported,
///    not silently ignored.
/// </summary>
public sealed class DocxTemplateRenderer
{
    private static readonly Regex PlaceholderRegex =
        new(@"\[\[(?<name>[A-Za-z0-9_]+(?:\.[A-Za-z0-9_]+)*)\]\]", RegexOptions.Compiled);

    private static readonly Regex RepeatStartRegex =
        new(@"^\[\[Repeat:(?<key>[A-Za-z0-9_]+)(?:\s+as\s+(?<alias>[A-Za-z0-9_]+))?\]\]$", RegexOptions.Compiled);

    private static readonly Regex RepeatEndRegex =
        new(@"^\[\[EndRepeat:(?<key>[A-Za-z0-9_]+)\]\]$", RegexOptions.Compiled);

    /// <summary>
    /// Renders the template and returns the generated document as bytes.
    /// Throws <see cref="TemplateValidationException"/> if validation fails; the
    /// source template bytes are never mutated (rendering happens on a copy).
    /// </summary>
    public byte[] Render(byte[] templateBytes, TemplateData data)
    {
        using var output = new MemoryStream();
        output.Write(templateBytes, 0, templateBytes.Length);
        output.Position = 0;

        using (var doc = WordprocessingDocument.Open(output, isEditable: true))
        {
            var mainPart = doc.MainDocumentPart
                ?? throw new InvalidOperationException("Template has no main document part.");

            var parts = new List<OpenXmlPart> { mainPart };
            parts.AddRange(mainPart.HeaderParts);
            parts.AddRange(mainPart.FooterParts);

            var errors = new List<string>();
            foreach (var part in parts)
                ProcessPart(part, data, errors);

            if (errors.Count > 0)
                throw new TemplateValidationException(errors);

            foreach (var part in parts)
                part.RootElement?.Save();
        }

        return output.ToArray();
    }

    // ---- Part-level orchestration -------------------------------------------------------

    private void ProcessPart(OpenXmlPart part, TemplateData data, List<string> errors)
    {
        var root = part.RootElement;
        if (root is null) return;

        // Block regions first — cloning changes the tree, and a block's own scalar/
        // collection placeholders are resolved recursively as each clone is created.
        ProcessBlocksRecursive(root, data, errors);

        // Whatever collections/scalars remain outside any block.
        ProcessSubtree(root, data, errors);
    }

    private void ProcessSubtree(OpenXmlElement root, TemplateData data, List<string> errors)
    {
        var tables = (root is Table t ? new[] { t } : Array.Empty<Table>())
            .Concat(root.Descendants<Table>()).ToList();
        foreach (var table in tables)
            ProcessTable(table, data, errors);

        var paragraphs = (root is Paragraph p ? new[] { p } : Array.Empty<Paragraph>())
            .Concat(root.Descendants<Paragraph>()).ToList();
        foreach (var paragraph in paragraphs)
            ReplaceScalarsInParagraph(paragraph, data, errors);
    }

    // ---- Block (Repeat/EndRepeat) handling ------------------------------------------------

    private void ProcessBlocksRecursive(OpenXmlElement container, TemplateData data, List<string> errors)
    {
        while (TryProcessOneBlock(container, data, errors)) { }

        foreach (var child in container.ChildElements.ToList())
        {
            if (child.HasChildren)
                ProcessBlocksRecursive(child, data, errors);
        }
    }

    private bool TryProcessOneBlock(OpenXmlElement container, TemplateData data, List<string> errors)
    {
        var children = container.ChildElements.ToList();

        var startIdx = -1;
        string? blockKey = null;
        string? alias = null;

        for (var i = 0; i < children.Count; i++)
        {
            if (children[i] is not Paragraph p) continue;
            var m = RepeatStartRegex.Match(GetParagraphText(p).Trim());
            if (!m.Success) continue;

            startIdx = i;
            blockKey = m.Groups["key"].Value;
            alias = m.Groups["alias"].Success ? m.Groups["alias"].Value : blockKey;
            break;
        }

        if (startIdx == -1) return false;

        var endIdx = -1;
        for (var i = startIdx + 1; i < children.Count; i++)
        {
            if (children[i] is not Paragraph p) continue;
            var m = RepeatEndRegex.Match(GetParagraphText(p).Trim());
            if (m.Success && m.Groups["key"].Value.Equals(blockKey, StringComparison.OrdinalIgnoreCase))
            {
                endIdx = i;
                break;
            }
        }

        var startMarker = children[startIdx];

        if (endIdx == -1)
        {
            errors.Add($"[[Repeat:{blockKey}]] has no matching [[EndRepeat:{blockKey}]]");
            startMarker.Remove(); // prevent an infinite loop on the next scan of this container
            return true;
        }

        var endMarker = children[endIdx];
        var contentRange = children.Skip(startIdx + 1).Take(endIdx - startIdx - 1).ToList();

        if (!data.Blocks.TryGetValue(blockKey!, out var items))
        {
            errors.Add($"Unknown block: {blockKey}");
            items = new List<TemplateData>();
        }

        OpenXmlElement anchor = endMarker;
        foreach (var item in items)
        {
            var effective = BuildEffectiveData(data, alias!, item);
            foreach (var node in contentRange)
            {
                var clone = node.CloneNode(deep: true);
                container.InsertAfter(clone, anchor);
                anchor = clone;

                ProcessBlocksRecursive(clone, effective, errors); // nested blocks, if any
                ProcessSubtree(clone, effective, errors);
            }
        }

        startMarker.Remove();
        endMarker.Remove();
        foreach (var node in contentRange) node.Remove();

        return true;
    }

    /// <summary>
    /// Builds the TemplateData used to resolve placeholders inside one clone of a block's
    /// content range: the parent scope's scalars/collections stay reachable unprefixed,
    /// plus the block item's own scalars/collections/blocks re-keyed under "{alias}.".
    /// </summary>
    private static TemplateData BuildEffectiveData(TemplateData parent, string alias, TemplateData item)
    {
        var effective = new TemplateData();

        foreach (var kv in parent.Scalars) effective.Scalars[kv.Key] = kv.Value;
        foreach (var kv in parent.Collections) effective.Collections[kv.Key] = kv.Value;
        foreach (var kv in parent.Blocks) effective.Blocks[kv.Key] = kv.Value;

        foreach (var kv in item.Scalars) effective.Scalars[$"{alias}.{kv.Key}"] = kv.Value;
        foreach (var kv in item.Collections) effective.Collections[$"{alias}.{kv.Key}"] = kv.Value;
        foreach (var kv in item.Blocks) effective.Blocks[$"{alias}.{kv.Key}"] = kv.Value;

        return effective;
    }

    private static string GetParagraphText(Paragraph p) => string.Concat(p.Descendants<Text>().Select(t => t.Text));

    // ---- Table / collection handling ----------------------------------------------------

    private void ProcessTable(Table table, TemplateData data, List<string> errors)
    {
        foreach (var row in table.Elements<TableRow>().ToList())
        {
            var rowText = string.Concat(row.Descendants<Text>().Select(t => t.Text));
            var placeholderNames = PlaceholderRegex.Matches(rowText)
                .Select(m => m.Groups["name"].Value)
                .ToList();

            var collectionRefs = placeholderNames
                .Select(p => MatchCollectionKey(p, data))
                .Where(k => k is not null)
                .Cast<string>()
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (collectionRefs.Count == 0)
                continue; // plain row, nothing to clone

            if (collectionRefs.Count > 1)
            {
                errors.Add($"A template row may reference only one collection. Found: {string.Join(", ", collectionRefs)}");
                continue;
            }

            var collectionName = collectionRefs[0];
            var items = data.Collections[collectionName];

            if (items.Count == 0)
            {
                row.Remove();
                continue;
            }

            var parent = row.Parent!;
            OpenXmlElement anchor = row;

            foreach (var item in items)
            {
                var clone = (TableRow)row.CloneNode(deep: true);
                ReplaceCollectionPlaceholders(clone, collectionName, item, errors);
                parent.InsertAfter(clone, anchor);
                anchor = clone;
            }

            row.Remove();
        }
    }

    /// <summary>
    /// Finds which known collection key a placeholder belongs to. Longest-match, since
    /// collection keys can themselves be dotted (e.g. "Card.Transactions" once inside
    /// a block), not just a single segment.
    /// </summary>
    private static string? MatchCollectionKey(string placeholderName, TemplateData data)
    {
        return data.Collections.Keys
            .Where(k => placeholderName.Equals(k, StringComparison.OrdinalIgnoreCase)
                     || placeholderName.StartsWith(k + ".", StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(k => k.Length)
            .FirstOrDefault();
    }

    private void ReplaceCollectionPlaceholders(
        TableRow row, string collectionName, Dictionary<string, string> item, List<string> errors)
    {
        foreach (var paragraph in row.Descendants<Paragraph>().ToList())
        {
            ReplaceInParagraph(paragraph, name =>
            {
                if (!name.StartsWith(collectionName + ".", StringComparison.OrdinalIgnoreCase))
                    return null; // belongs to a different placeholder kind — leave for another pass

                var propertyName = name[(collectionName.Length + 1)..];
                if (!item.TryGetValue(propertyName, out var value))
                {
                    errors.Add($"Unknown placeholder: {name}");
                    return string.Empty;
                }

                return value;
            });
        }
    }

    // ---- Scalar handling ------------------------------------------------------------------

    private void ReplaceScalarsInParagraph(Paragraph paragraph, TemplateData data, List<string> errors)
    {
        ReplaceInParagraph(paragraph, name =>
        {
            // Belongs to a collection row — already handled (or reported) during table processing.
            if (MatchCollectionKey(name, data) is not null)
                return null;

            if (!data.Scalars.TryGetValue(name, out var value))
            {
                errors.Add($"Unknown placeholder: {name}");
                return string.Empty;
            }

            return value;
        });
    }

    // ---- Core run-aware find & replace -----------------------------------------------------

    /// <summary>
    /// Finds every [[Name]] placeholder in a paragraph's visible text — reassembled across
    /// however many runs Word split it into — and replaces it using <paramref name="resolve"/>.
    /// Returning null from <paramref name="resolve"/> leaves that placeholder untouched, so the
    /// same paragraph can be safely visited by multiple passes.
    /// </summary>
    private void ReplaceInParagraph(Paragraph paragraph, Func<string, string?> resolve)
    {
        var textNodes = paragraph.Descendants<Text>().ToList();
        if (textNodes.Count == 0) return;

        var sb = new StringBuilder();
        var spans = new List<(Text Node, int Start, int Length)>();
        foreach (var node in textNodes)
        {
            var start = sb.Length;
            sb.Append(node.Text);
            spans.Add((node, start, node.Text.Length));
        }

        var fullText = sb.ToString();
        var matches = PlaceholderRegex.Matches(fullText);
        if (matches.Count == 0) return;

        foreach (Match match in matches.Cast<Match>().Reverse())
        {
            var name = match.Groups["name"].Value;
            var replacement = resolve(name);
            if (replacement is null) continue;

            var matchStart = match.Index;
            var matchEnd = match.Index + match.Length;

            var affected = spans.Where(s => s.Start + s.Length > matchStart && s.Start < matchEnd).ToList();
            if (affected.Count == 0) continue;

            for (var i = 0; i < affected.Count; i++)
            {
                var (node, start, length) = affected[i];
                var nodeText = node.Text;
                var isFirst = i == 0;
                var isLast = i == affected.Count - 1;

                var localStart = Math.Max(matchStart - start, 0);
                var localEnd = Math.Min(matchEnd - start, length);

                var prefix = isFirst ? nodeText[..localStart] : string.Empty;
                var suffix = isLast ? nodeText[localEnd..] : string.Empty;
                var insert = isFirst ? replacement : string.Empty;

                node.Text = prefix + insert + suffix;
                node.Space = SpaceProcessingModeValues.Preserve;
            }
        }
    }
}
