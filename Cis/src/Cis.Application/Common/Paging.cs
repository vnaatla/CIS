using System.Collections;
using System.Globalization;
using System.Reflection;

namespace Cis.Application.Common;

/// <summary>One page of a search and the number of matching rows (server paging of the shared grid).</summary>
/// <typeparam name="T">Row type.</typeparam>
/// <param name="Rows">The page.</param>
/// <param name="Total">All matching rows.</param>
public record PagedResult<T>(List<T> Rows, int Total);

/// <summary>Grid paging request: 0-based page, rows per page, optional sort column (a row key / property name) and direction.</summary>
/// <param name="Page">0-based page.</param>
/// <param name="PageSize">Rows per page (1–1,000).</param>
/// <param name="SortKey">Column to sort by (null = the search's own order).</param>
/// <param name="SortDesc">Descending.</param>
public record PageQuery(int Page = 0, int PageSize = 50, string? SortKey = null, bool SortDesc = false)
{
    /// <summary>Largest page the API serves.</summary>
    public const int MaxPageSize = 1000;
}

/// <summary>Sorts and cuts an in-memory search result (searches of a few thousand rows; very large ones page in SQL).</summary>
public static class Paging
{
    /// <summary>The requested page of <paramref name="rows"/>, sorted by <see cref="PageQuery.SortKey"/> when given.</summary>
    public static PagedResult<T> Apply<T>(IReadOnlyList<T> rows, PageQuery q)
    {
        var size = Math.Clamp(q.PageSize, 1, PageQuery.MaxPageSize);
        IEnumerable<T> sorted = rows;
        if (!string.IsNullOrWhiteSpace(q.SortKey))
        {
            var key = q.SortKey.Trim();
            sorted = q.SortDesc ? rows.OrderByDescending(r => Value(r, key), Comparer) : rows.OrderBy(r => Value(r, key), Comparer);
        }
        return new PagedResult<T>(sorted.Skip(Math.Max(0, q.Page) * size).Take(size).ToList(), rows.Count);
    }

    private static readonly IComparer<object?> Comparer = Comparer<object?>.Create(Compare);

    /// <summary>Nulls first; numbers and dates by value; text case-insensitively.</summary>
    private static int Compare(object? a, object? b)
    {
        if (a is null || b is null) return a is null ? (b is null ? 0 : -1) : 1;
        if (IsNumber(a) && IsNumber(b)) return Convert.ToDecimal(a, CultureInfo.InvariantCulture).CompareTo(Convert.ToDecimal(b, CultureInfo.InvariantCulture));
        if (a is DateTime da && b is DateTime db) return da.CompareTo(db);
        return string.Compare(Convert.ToString(a, CultureInfo.InvariantCulture)?.Trim(), Convert.ToString(b, CultureInfo.InvariantCulture)?.Trim(),
            StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsNumber(object v) => v is byte or short or int or long or decimal or double or float;

    /// <summary>A dictionary row's value (key case-insensitive) or a record's property (name case-insensitive).</summary>
    private static object? Value<T>(T row, string key)
    {
        if (row is IDictionary dict)
        {
            foreach (DictionaryEntry e in dict)
                if (string.Equals(e.Key as string, key, StringComparison.OrdinalIgnoreCase)) return e.Value;
            return null;
        }
        return typeof(T).GetProperty(key, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase)?.GetValue(row);
    }
}
