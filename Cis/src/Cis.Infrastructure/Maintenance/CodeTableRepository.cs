using System.Text;
using System.Text.RegularExpressions;
using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.CodeTables;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Maintenance;

/// <summary>
/// <see cref="ICodeTableRepository"/>. The simple maintenance tables share one implementation instead of five procedures each: the SQL is
/// built only from the identifiers of <see cref="CodeTableCatalog"/> (compile-time constants, checked against <see cref="SafeName"/> and
/// bracket-quoted), and every value is a parameter. Nothing from the request becomes SQL text.
/// </summary>
public sealed partial class CodeTableRepository : ICodeTableRepository
{
    /// <summary>Allowed shape of a catalog table / column name.</summary>
    [GeneratedRegex("^[A-Za-z_][A-Za-z0-9_]*$")]
    public static partial Regex SafeName();

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, CodeTableDefinition table, IReadOnlyDictionary<string, string> filters, int max)
    {
        var sql = new StringBuilder($"SELECT TOP (@max) {string.Join(", ", table.Fields.Select(f => Q(f.Column)))} FROM dbo.{Q(table.Table)} WHERE 1 = 1");
        var p = new DynamicParameters(new { max });
        var i = 0;
        foreach (var field in table.Fields.Where(f => filters.ContainsKey(f.Column)))
        {
            var name = $"f{i++}";
            var value = filters[field.Column];
            if (field.Search == SearchMatch.Contains)
            {
                sql.Append($" AND {Q(field.Column)} LIKE @{name}");
                p.Add(name, "%" + SqlText.EscapeLike(value) + "%");
            }
            else
            {
                sql.Append($" AND {Q(field.Column)} = @{name}");
                p.Add(name, value);
            }
        }
        sql.Append($" ORDER BY {Q(table.KeyField.Column)}");
        var rows = await session.Sql().Connection.QueryAsync(sql.ToString(), p, session.Sql().Transaction);
        return RowMapper.ToDictionaries(rows);
    }

    /// <inheritdoc />
    public async Task<bool> ExistsAsync(IDbSession session, CodeTableDefinition table, string code) =>
        await session.Sql().Connection.ExecuteScalarAsync<int>(
            $"SELECT CASE WHEN EXISTS (SELECT 1 FROM dbo.{Q(table.Table)} WHERE {Q(table.KeyField.Column)} = @code) THEN 1 ELSE 0 END",
            new { code }, session.Sql().Transaction) == 1;

    /// <inheritdoc />
    public Task InsertAsync(IDbSession session, CodeTableDefinition table, IReadOnlyDictionary<string, string?> values, string userCode, DateTime date)
    {
        var (columns, p) = Assignments(table, values, userCode, date, includeKey: !table.IdentityKey);
        var sql = $"INSERT INTO dbo.{Q(table.Table)} ({string.Join(", ", columns.Select(Q))}) VALUES ({string.Join(", ", columns.Select((_, n) => "@p" + n))})";
        return session.Sql().Connection.ExecuteAsync(sql, p, session.Sql().Transaction);
    }

    /// <inheritdoc />
    public Task<int> UpdateFirstAsync(IDbSession session, CodeTableDefinition table, string code, IReadOnlyDictionary<string, string?> values, string userCode, DateTime date)
    {
        var (columns, p) = Assignments(table, values, userCode, date, includeKey: false);
        var sql = $"UPDATE TOP (1) dbo.{Q(table.Table)} SET {string.Join(", ", columns.Select((c, n) => $"{Q(c)} = @p{n}"))} WHERE {Q(table.KeyField.Column)} = @code";
        p.Add("code", code);
        return session.Sql().Connection.ExecuteAsync(sql, p, session.Sql().Transaction);
    }

    /// <inheritdoc />
    public Task<int> DeleteAsync(IDbSession session, CodeTableDefinition table, string code) =>
        session.Sql().Connection.ExecuteAsync($"DELETE FROM dbo.{Q(table.Table)} WHERE {Q(table.KeyField.Column)} = @code", new { code }, session.Sql().Transaction);

    /// <inheritdoc />
    public async Task<List<LookupOption>> ChoicesAsync(IDbSession session, ChoiceSource source) =>
        (await session.Sql().Connection.QueryAsync<LookupOption>(
            $"SELECT RTRIM({Q(source.CodeColumn)}) AS Code, ISNULL(RTRIM({Q(source.NameColumn)}), N'') AS Name FROM dbo.{Q(source.Table)} ORDER BY {Q(source.CodeColumn)}",
            transaction: session.Sql().Transaction)).ToList();

    /// <summary>Columns written and their values (@p0..): the fields (key only when asked), then the audit columns the table has.</summary>
    private static (List<string> Columns, DynamicParameters Parameters) Assignments(CodeTableDefinition table, IReadOnlyDictionary<string, string?> values,
        string userCode, DateTime date, bool includeKey)
    {
        var columns = new List<string>();
        var p = new DynamicParameters();
        void Add(string column, object? value)
        {
            p.Add("p" + columns.Count, value);
            columns.Add(column);
        }
        foreach (var field in table.Fields.Where(f => includeKey || !f.IsKey)) Add(field.Column, values.TryGetValue(field.Column, out var v) ? v : null);
        if (table.AuditUserColumn is not null) Add(table.AuditUserColumn, userCode);
        if (table.AuditDateColumn is not null) Add(table.AuditDateColumn, date);
        return (columns, p);
    }

    private static string Q(string name) =>
        SafeName().IsMatch(name) ? "[" + name + "]" : throw new InvalidOperationException($"Unsafe identifier '{name}' in the code table catalog.");
}
