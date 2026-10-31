using Cis.Application.Abstractions;
using Microsoft.Data.SqlClient;

namespace Cis.Infrastructure.Data;

/// <summary>An <see cref="IDbSession"/> on SQL Server: the connection and transaction repositories run their commands on.</summary>
public interface ISqlSession : IDbSession
{
    /// <summary>The open connection (CISDB).</summary>
    SqlConnection Connection { get; }

    /// <summary>The open transaction, or null for a read-only session.</summary>
    SqlTransaction? Transaction { get; }
}

/// <summary><see cref="IDbSessions"/> on the main database (CISDB).</summary>
public sealed class SqlDbSessions(DbConnections db) : IDbSessions
{
    /// <inheritdoc />
    public async Task<IDbSession> OpenAsync()
    {
        var connection = db.Main();
        await connection.OpenAsync();
        return new SqlSession(connection, null);
    }

    /// <inheritdoc />
    public async Task<IDbSession> BeginAsync()
    {
        var connection = db.Main();
        await connection.OpenAsync();
        return new SqlSession(connection, (SqlTransaction)await connection.BeginTransactionAsync());
    }

    private sealed class SqlSession(SqlConnection connection, SqlTransaction? transaction) : ISqlSession
    {
        private bool _finished;

        public SqlConnection Connection => connection;
        public SqlTransaction? Transaction => transaction;
        public bool IsTransactional => transaction is not null;

        public async Task CommitAsync()
        {
            if (transaction is null || _finished) return;
            _finished = true;
            await transaction.CommitAsync();
        }

        public async Task RollbackAsync()
        {
            if (transaction is null || _finished) return;
            _finished = true;
            await transaction.RollbackAsync();
        }

        public async ValueTask DisposeAsync()
        {
            try
            {
                await RollbackAsync();
            }
            catch (InvalidOperationException)
            {
                // The connection already failed; the server has rolled the transaction back and the original error is the one to report.
            }
            finally
            {
                if (transaction is not null) await transaction.DisposeAsync();
                await connection.DisposeAsync();
            }
        }
    }
}
