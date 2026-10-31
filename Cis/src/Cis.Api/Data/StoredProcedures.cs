using System.Data;
using Dapper;

namespace Cis.Api.Data;

/// <summary>Calls to the <c>cis.*</c> stored procedures (scripts in sql/main) on a session. Guarded writes return a Result column.</summary>
public static class StoredProcedures
{
    /// <summary>The write was done.</summary>
    public const int Ok = 0;

    /// <summary>The record asked for does not exist.</summary>
    public const int NotFound = 1;

    /// <summary>The write was refused: the record exists already or is still in use.</summary>
    public const int Conflict = 2;

    /// <summary>Executes a procedure and returns the rows affected.</summary>
    public static Task<int> ExecAsync(this IDbSession s, string name, object? p = null) =>
        s.Connection.ExecuteAsync(name, p, s.Transaction, commandType: CommandType.StoredProcedure);

    /// <summary>The first row of a procedure's result, or default.</summary>
    public static Task<T?> FirstOrDefaultAsync<T>(this IDbSession s, string name, object? p = null) =>
        s.Connection.QueryFirstOrDefaultAsync<T>(name, p, s.Transaction, commandType: CommandType.StoredProcedure);

    /// <summary>The first column of the first row of a procedure's result.</summary>
    public static Task<T?> ScalarAsync<T>(this IDbSession s, string name, object? p = null) =>
        s.Connection.ExecuteScalarAsync<T>(name, p, s.Transaction, commandType: CommandType.StoredProcedure);

    /// <summary>Every row of a procedure's result.</summary>
    public static async Task<List<T>> QueryAsync<T>(this IDbSession s, string name, object? p = null) =>
        (await s.Connection.QueryAsync<T>(name, p, s.Transaction, commandType: CommandType.StoredProcedure)).ToList();

    /// <summary>A procedure that returns several result sets; dispose the reader when done.</summary>
    public static Task<SqlMapper.GridReader> MultipleAsync(this IDbSession s, string name, object? p = null) =>
        s.Connection.QueryMultipleAsync(name, p, s.Transaction, commandType: CommandType.StoredProcedure);

    /// <summary>Rows of a procedure as case-insensitive dictionaries (legacy procedures with many columns, report queries).</summary>
    public static async Task<List<Dictionary<string, object?>>> RowsAsync(this IDbSession s, string name, object? p = null) =>
        RowMapper.ToDictionaries(await s.Connection.QueryAsync(name, p, s.Transaction, commandType: CommandType.StoredProcedure));
}
