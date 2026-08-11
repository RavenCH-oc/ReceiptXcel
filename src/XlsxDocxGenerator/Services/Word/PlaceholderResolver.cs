using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Word;

/// <summary>
/// Finds and replaces placeholders across Text nodes in a paragraph.
/// The first Text node touched by a placeholder receives the replacement value;
/// prefix/suffix text stays in its original runs, so surrounding formatting is preserved.
/// </summary>
public interface IPlaceholderResolver
{
    IReadOnlyList<PlaceholderOccurrence> FindOccurrences(OpenXmlElement root);

    void ReplaceOccurrences(
        OpenXmlElement root,
        IReadOnlyDictionary<string, string?> values);
}

public sealed record PlaceholderOccurrence(
    string Name,
    IReadOnlyList<Text> TextNodes);

public sealed class PlaceholderResolver : IPlaceholderResolver
{
    private static readonly Regex PlaceholderPattern = new(
        "\\{\\{(?<name>[A-Za-z_][A-Za-z0-9_]*)\\}\\}",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public IReadOnlyList<PlaceholderOccurrence> FindOccurrences(OpenXmlElement root)
    {
        ArgumentNullException.ThrowIfNull(root);

        return FindLocatedOccurrences(root)
            .Select(occurrence => new PlaceholderOccurrence(
                occurrence.Name,
                occurrence.TextNodes.Select(node => node.Node).ToArray()))
            .ToArray();
    }

    public void ReplaceOccurrences(
        OpenXmlElement root,
        IReadOnlyDictionary<string, string?> values)
    {
        ArgumentNullException.ThrowIfNull(root);
        ArgumentNullException.ThrowIfNull(values);

        var canonicalValues = values
            .ToDictionary(
                pair => pair.Key.ToUpperInvariant(),
                pair => pair.Value,
                StringComparer.OrdinalIgnoreCase);

        foreach (var paragraph in root.Descendants<Paragraph>().ToArray())
        {
            var matches = FindLocatedOccurrences(paragraph)
                .OrderByDescending(occurrence => occurrence.Start)
                .ToArray();

            foreach (var occurrence in matches)
            {
                if (!canonicalValues.TryGetValue(occurrence.Name, out var replacement))
                {
                    throw new GenerationException(
                        GenerationErrorCode.PlaceholderReplacementFailed,
                        $"找不到 placeholder「{occurrence.Name}」的替換值。");
                }

                ReplaceOne(occurrence, replacement ?? string.Empty);
            }
        }
    }

    private static IReadOnlyList<LocatedOccurrence> FindLocatedOccurrences(OpenXmlElement root)
    {
        return root.Descendants<Paragraph>()
            .SelectMany(FindLocatedOccurrences)
            .ToArray();
    }

    private static IReadOnlyList<LocatedOccurrence> FindLocatedOccurrences(Paragraph paragraph)
    {
        var textNodes = paragraph.Descendants<Text>().ToArray();
        if (textNodes.Length == 0)
        {
            return [];
        }

        var nodeSpans = new List<TextNodeSpan>(textNodes.Length);
        var combinedText = string.Empty;
        foreach (var textNode in textNodes)
        {
            var value = textNode.Text ?? string.Empty;
            var start = combinedText.Length;
            combinedText += value;
            nodeSpans.Add(new TextNodeSpan(textNode, start, combinedText.Length));
        }

        return PlaceholderPattern.Matches(combinedText)
            .Cast<Match>()
            .Select(match =>
            {
                var end = match.Index + match.Length;
                var involvedNodes = nodeSpans
                    .Where(span => span.Start < end && span.End > match.Index)
                    .ToArray();

                return new LocatedOccurrence(
                    match.Groups["name"].Value.ToUpperInvariant(),
                    match.Index,
                    end,
                    involvedNodes);
            })
            .ToArray();
    }

    private static void ReplaceOne(LocatedOccurrence occurrence, string replacement)
    {
        var firstNode = occurrence.TextNodes[0];
        var lastNode = occurrence.TextNodes[^1];
        var firstValue = firstNode.Node.Text ?? string.Empty;
        var lastValue = lastNode.Node.Text ?? string.Empty;

        var firstLocalStart = occurrence.Start - firstNode.Start;
        var lastLocalEnd = occurrence.End - lastNode.Start;

        if (ReferenceEquals(firstNode.Node, lastNode.Node))
        {
            SetText(firstNode.Node,
                firstValue[..firstLocalStart]
                + replacement
                + firstValue[lastLocalEnd..]);
            return;
        }

        SetText(firstNode.Node, firstValue[..firstLocalStart] + replacement);

        foreach (var middleNode in occurrence.TextNodes.Skip(1).SkipLast(1))
        {
            SetText(middleNode.Node, string.Empty);
        }

        SetText(lastNode.Node, lastValue[lastLocalEnd..]);
    }

    private static void SetText(Text textNode, string value)
    {
        textNode.Text = value;
        textNode.Space = value.Length > 0
            && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]))
            ? SpaceProcessingModeValues.Preserve
            : null;
    }

    private sealed record TextNodeSpan(Text Node, int Start, int End);

    private sealed record LocatedOccurrence(
        string Name,
        int Start,
        int End,
        IReadOnlyList<TextNodeSpan> TextNodes);
}
