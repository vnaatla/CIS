using System.Data;
using Cis.Application.Abstractions;
using Dapper;

namespace Cis.Infrastructure.Data;

/// <summary>Calls to the <c>cis.*</c> stored procedures (scripts in sql/main) on a session. Guarded writes return a Result column.</summary>
public static class StoredProcedures
{
    /// <summary>The write was done.</summary>
    public const int Ok = 0;

    /// <summary>The record asked for does not exist.</summary>
    public const int NotFound = 1;

    /// <summary>The write was refused: the record exists already or is still in use.</summary>
    public const int Conflict = 2;

    /// <summary>The SQL connection behind an application session (sessions come from <see cref="SqlDbSessions"/>).</summary>
    public static ISqlSession Sql(this IDbSession session) =>
        session as ISqlSession ?? throw new InvalidOperationException($"Repositories need a session created by {nameof(SqlDbSessions)}.");

    /// <summary>Executes a procedure and returns the rows affected.</summary>
    public static Task<int> ExecAsync(this IDbSession s, string name, object? p = null) =>
        s.Sql().Connection.ExecuteAsync(name, p, s.Sql().Transaction, commandType: CommandType.StoredProcedure);

    /// <summary>The first row of a procedure's result, or default.</summary>
    public static Task<T?> FirstOrDefaultAsync<T>(this IDbSession s, string name, object? p = null) =>
        s.Sql().Connection.QueryFirstOrDefaultAsync<T>(name, p, s.Sql().Transaction, commandType: CommandType.StoredProcedure);

    /// <summary>The first column of the first row of a procedure's result.</summary>
    public static Task<T?> ScalarAsync<T>(this IDbSession s, string name, object? p = null) =>
        s.Sql().Connection.ExecuteScalarAsync<T>(name, p, s.Sql().Transaction, commandType: CommandType.StoredProcedure);

    /// <summary>Every row of a procedure's result.</summary>
    public static async Task<List<T>> QueryAsync<T>(this IDbSession s, string name, object? p = null) =>
        (await s.Sql().Connection.QueryAsync<T>(name, p, s.Sql().Transaction, commandType: CommandType.StoredProcedure)).ToList();

    /// <summary>A procedure that returns several result sets; dispose the reader when done.</summary>
    public static Task<SqlMapper.GridReader> MultipleAsync(this IDbSession s, string name, object? p = null) =>
        s.Sql().Connection.QueryMultipleAsync(name, p, s.Sql().Transaction, commandType: CommandType.StoredProcedure);

    /// <summary>Rows of a procedure as case-insensitive dictionaries (legacy procedures with many columns, report queries).</summary>
    public static async Task<List<Dictionary<string, object?>>> RowsAsync(this IDbSession s, string name, object? p = null) =>
        RowMapper.ToDictionaries(await s.Sql().Connection.QueryAsync(name, p, s.Sql().Transaction, commandType: CommandType.StoredProcedure));
}
