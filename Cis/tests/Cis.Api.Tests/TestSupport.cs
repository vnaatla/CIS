using Cis.Application.Abstractions;
using Cis.Infrastructure.Data;
using FluentValidation;
using Microsoft.Extensions.Configuration;

namespace Cis.Api.Tests;

/// <summary>A unit of work that touches no database: services can be tested against fake repositories.</summary>
internal sealed class FakeSession(bool transactional) : IDbSession
{
    public bool Committed { get; private set; }
    public bool RolledBack { get; private set; }
    public bool IsTransactional => transactional;

    public Task CommitAsync()
    {
        Committed = true;
        return Task.CompletedTask;
    }

    public Task RollbackAsync()
    {
        RolledBack = !Committed;
        return Task.CompletedTask;
    }

    public async ValueTask DisposeAsync() => await RollbackAsync();
}

/// <summary>Hands out <see cref="FakeSession"/>s and remembers them.</summary>
internal sealed class FakeSessions : IDbSessions
{
    public List<FakeSession> Sessions { get; } = [];
    public FakeSession Last => Sessions[^1];
    public Task<IDbSession> OpenAsync() => Task.FromResult(Add(new FakeSession(false)));
    public Task<IDbSession> BeginAsync() => Task.FromResult(Add(new FakeSession(true)));

    private IDbSession Add(FakeSession session)
    {
        Sessions.Add(session);
        return session;
    }
}

/// <summary>A clock stopped at a given local time.</summary>
internal sealed class FixedClock(DateTime local) : TimeProvider
{
    public override TimeZoneInfo LocalTimeZone => TimeZoneInfo.Utc;
    public override DateTimeOffset GetUtcNow() => new(DateTime.SpecifyKind(local, DateTimeKind.Utc));
}

/// <summary>Connection of the database tests (skipped when CIS_TEST_DB is not set).</summary>
internal static class TestDb
{
    public static readonly string? Cis = Environment.GetEnvironmentVariable("CIS_TEST_DB");

    public static DbConnections Connections() => new(new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
    {
        [$"ConnectionStrings:{DbConnections.MainName}"] = Cis,
    }).Build());

    public static IDbSessions Sessions() => new SqlDbSessions(Connections());
}

internal static class ValidatorExtensions
{
    /// <summary>The first message the validator reports (what the API answers with), or null when the request is valid.</summary>
    public static string? FirstError<T>(this IValidator<T> validator, T request) => validator.Validate(request).Errors.FirstOrDefault()?.ErrorMessage;
}

/// <summary>Folders of the repository (D:\CIS).</summary>
internal static class RepoPaths
{
    public static readonly string CisRoot = FindUp(d => Directory.Exists(Path.Combine(d, "sql", "main")));
    public static readonly string RepoRoot = Directory.GetParent(CisRoot)!.FullName;

    private static string FindUp(Func<string, bool> match)
    {
        for (var d = new DirectoryInfo(AppContext.BaseDirectory); d is not null; d = d.Parent)
            if (match(d.FullName)) return d.FullName;
        throw new DirectoryNotFoundException("Cis root (sql/main) not found above " + AppContext.BaseDirectory);
    }
}
