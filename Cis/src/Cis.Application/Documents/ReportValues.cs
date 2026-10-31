using System.Globalization;

namespace Cis.Application.Documents;

/// <summary>Turns report cell values into numbers / text the same way for every writer.</summary>
public static class ReportValues
{
    /// <summary>The raw value of <paramref name="column"/> in <paramref name="row"/>.</summary>
    public static object? Get(IReadOnlyDictionary<string, object?> row, ReportColumn column) =>
        row.TryGetValue(column.Key, out var v) ? v : row.FirstOrDefault(kv => string.Equals(kv.Key, column.Key, StringComparison.OrdinalIgnoreCase)).Value;

    /// <summary>The value as a number, or null when it is not numeric.</summary>
    public static decimal? Number(object? value) => value switch
    {
        null or DBNull => null,
        decimal d => d,
        double d => (decimal)d,
        float f => (decimal)f,
        int i => i,
        long l => l,
        short s => s,
        byte b => b,
        string s when decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) => d,
        _ => null,
    };

    /// <summary>The value as a date, or null.</summary>
    public static DateTime? Date(object? value) => value switch
    {
        DateTime d => d,
        DateTimeOffset o => o.DateTime,
        string s when DateTime.TryParseExact(s, ["dd/MM/yyyy", "yyyy-MM-dd", "yyyy-MM-ddTHH:mm:ss"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) => d,
        _ => null,
    };

    /// <summary>Display text of a cell (dd/mm/yyyy dates, 2-decimal money).</summary>
    public static string Text(object? value, ReportFormat format) => format switch
    {
        ReportFormat.Money => Number(value)?.ToString("#,##0.00", CultureInfo.InvariantCulture) ?? Plain(value),
        ReportFormat.Integer => Number(value)?.ToString("0", CultureInfo.InvariantCulture) ?? Plain(value),
        ReportFormat.Number => Number(value)?.ToString(CultureInfo.InvariantCulture) ?? Plain(value),
        ReportFormat.Date => Date(value)?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) ?? Plain(value),
        _ => Plain(value),
    };

    /// <summary>Sum of a totals column.</summary>
    public static decimal Sum(ReportDocument report, ReportColumn column) => report.Rows.Sum(r => Number(Get(r, column)) ?? 0m);

    private static string Plain(object? value) => value switch
    {
        null or DBNull => "",
        string s => s.Trim(),
        DateTime d => d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
        bool b => b ? "Yes" : "No",
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
        _ => value.ToString() ?? "",
    };
}
