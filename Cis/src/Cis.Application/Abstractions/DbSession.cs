namespace Cis.Application.Abstractions;

/// <summary>
/// One unit of work on the database: for writes, a transaction. Services decide where it begins and ends and hand it to repositories;
/// repositories never commit. Disposing a session that was not committed rolls it back, so an exception can never leave half a save.
/// The application layer does not know it is SQL Server: the infrastructure's session carries the connection.
/// </summary>
public interface IDbSession : IAsyncDisposable
{
    /// <summary>True when the session holds a transaction (<see cref="IDbSessions.BeginAsync"/>).</summary>
    bool IsTransactional { get; }

    /// <summary>Commits the transaction (no-op for a read-only session).</summary>
    Task CommitAsync();

    /// <summary>Rolls the transaction back (no-op for a read-only session or once finished).</summary>
    Task RollbackAsync();
}

/// <summary>Creates <see cref="IDbSession"/>s on the main database (CISDB; the other databases are reached through synonyms).</summary>
public interface IDbSessions
{
    /// <summary>A read-only session.</summary>
    Task<IDbSession> OpenAsync();

    /// <summary>A session with a transaction; rolled back on dispose unless <see cref="IDbSession.CommitAsync"/> was called.</summary>
    Task<IDbSession> BeginAsync();
}
