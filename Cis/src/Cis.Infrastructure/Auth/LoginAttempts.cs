using Cis.Application.Auth.Login;
using Microsoft.Extensions.Caching.Memory;

namespace Cis.Infrastructure.Auth;

/// <summary><see cref="ILoginAttempts"/> in the API process memory, per user code and client address, for <see cref="Window"/>.</summary>
public sealed class LoginAttempts(IMemoryCache cache) : ILoginAttempts
{
    /// <summary>How long failures are remembered.</summary>
    public static readonly TimeSpan Window = TimeSpan.FromMinutes(30);

    /// <inheritdoc />
    public int Failures(string userCode, string? client) => cache.TryGetValue(Key(userCode, client), out int n) ? n : 0;

    /// <inheritdoc />
    public void RegisterFailure(string userCode, string? client) =>
        cache.Set(Key(userCode, client), Failures(userCode, client) + 1, Window);

    /// <inheritdoc />
    public void Reset(string userCode, string? client) => cache.Remove(Key(userCode, client));

    private static string Key(string userCode, string? client) => $"login:{userCode.Trim().ToUpperInvariant()}:{client}";
}
