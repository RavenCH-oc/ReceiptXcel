namespace XlsxDocxGenerator.Models;

public enum SelectionMode
{
    Marker,
    ExcelRows,
    Latest
}

/// <summary>
/// Validated, typed selection rule. Raw UI text is parsed before reaching RowResolver.
/// </summary>
public abstract record SelectionRule
{
    public sealed record Marker(
        string SelectionMarkerColumnHeader,
        IReadOnlyList<string> Tokens) : SelectionRule;

    public sealed record ExcelRows(IReadOnlyList<int> RowNumbers) : SelectionRule;

    public sealed record Latest(int Count) : SelectionRule;
}
