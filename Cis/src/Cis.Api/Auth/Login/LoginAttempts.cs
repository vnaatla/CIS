using Microsoft.Extensions.Caching.Memory;

namespace Cis.Api.Auth.Login;

/// <summary>
/// Wrong-password counter of frmLogin (<c>counter</c>, <c>txtCheckLogon</c>). The desktop kept it in the open login window; the web
/// keeps it per user code and client address for <see cref="LoginAttempts.Window"/>, which is the closest stateless equivalent.
/// </summary>
public interface ILoginAttempts
{
    /// <summary>Wrong passwords recorded for the user from this client.</summary>
    int Failures(string userCode, string? client);

    /// <summary>Adds a wrong password.</summary>
    void RegisterFailure(string userCode, string? client);

    /// <summary>Forgets the failures (successful login, or the account was suspended).</summary>
    void Reset(string userCode, string? client);
}

/// <inheritdoc />
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
