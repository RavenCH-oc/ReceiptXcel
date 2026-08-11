using System.Text.RegularExpressions;
using XlsxDocxGenerator.Models;
using XlsxDocxGenerator.Services.Errors;

namespace XlsxDocxGenerator.Services.Excel;

public sealed class SelectionRuleParser
{
    private static readonly Regex RowTokenPattern = new(
        "^(?<start>[0-9]+)(?:\\s*-\\s*(?<end>[0-9]+))?$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public SelectionRule Parse(
        SelectionMode mode,
        string? parameter,
        string? markerColumnHeader = null)
    {
        return mode switch
        {
            SelectionMode.Marker => ParseMarker(markerColumnHeader, parameter),
            SelectionMode.ExcelRows => ParseExcelRows(parameter),
            SelectionMode.Latest => ParseLatest(parameter),
            _ => throw new SelectionValidationException(
                SelectionErrorCode.InvalidSyntax,
                "不支援的資料選擇方式。")
        };
    }

    public SelectionRule.Marker ParseMarker(string? markerColumnHeader, string? parameter)
    {
        if (string.IsNullOrWhiteSpace(markerColumnHeader))
        {
            throw new SelectionValidationException(
                SelectionErrorCode.MissingMarkerColumn,
                "標記代號模式需要指定標記欄位。");
        }

        var tokens = SplitTokens(parameter);
        if (tokens.Count == 0)
        {
            throw new SelectionValidationException(
                SelectionErrorCode.MissingMarkerToken,
                "請輸入至少一個標記代號。");
        }

        return new SelectionRule.Marker(markerColumnHeader.Trim(), tokens);
    }

    public SelectionRule.ExcelRows ParseExcelRows(string? parameter)
    {
        if (string.IsNullOrWhiteSpace(parameter))
        {
            throw new SelectionValidationException(
                SelectionErrorCode.InvalidSyntax,
                "Excel 列號不可為空白。");
        }

        var rows = new SortedSet<int>();
        foreach (var token in parameter.Split(',', StringSplitOptions.TrimEntries))
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new SelectionValidationException(
                    SelectionErrorCode.InvalidSyntax,
                    "Excel 列號格式包含空白項目。");
            }

            var match = RowTokenPattern.Match(token);
            if (!match.Success
                || !int.TryParse(match.Groups["start"].Value, out var start)
                || start <= 0)
            {
                throw new SelectionValidationException(
                    SelectionErrorCode.InvalidSyntax,
                    $"無法解析 Excel 列號：{token}");
            }

            var end = start;
            if (match.Groups["end"].Success
                && (!int.TryParse(match.Groups["end"].Value, out end) || end <= 0))
            {
                throw new SelectionValidationException(
                    SelectionErrorCode.InvalidSyntax,
                    $"無法解析 Excel 列號範圍：{token}");
            }

            if (end < start)
            {
                throw new SelectionValidationException(
                    SelectionErrorCode.InvalidSyntax,
                    $"Excel 列號範圍不可反向：{token}");
            }

            for (var row = start; row <= end; row++)
            {
                rows.Add(row);
                if (row == int.MaxValue)
                {
                    break;
                }
            }
        }

        return new SelectionRule.ExcelRows(rows.ToArray());
    }

    public SelectionRule.Latest ParseLatest(string? parameter)
    {
        if (!int.TryParse(parameter?.Trim(), out var count) || count <= 0)
        {
            throw new SelectionValidationException(
                SelectionErrorCode.InvalidLatestCount,
                "最新資料筆數必須是大於 0 的整數。");
        }

        return new SelectionRule.Latest(count);
    }

    private static IReadOnlyList<string> SplitTokens(string? value) =>
        (value ?? string.Empty)
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
}
