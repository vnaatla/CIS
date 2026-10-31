using System.Globalization;
using Cis.Application.Auth.Login;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace Cis.Infrastructure.Auth;

/// <summary>Settings of the session tracking (configuration section <c>Session</c>).</summary>
public class SessionTrackingOptions
{
    /// <summary>Configuration section name.</summary>
    public const string SectionName = "Session";

    /// <summary>
    /// The desktop's <c>DocPath\FileLogin\LoginCIS.txt</c>. Empty = not written. Point it at the same share as the desktop to keep one
    /// log for both.
    /// </summary>
    public string? LoginLogPath { get; set; }
}

/// <summary>
/// <see cref="ILoginLog"/> that writes the line exactly as VB's <c>Write #</c> did: <c>"Login","USER",#yyyy-MM-dd#,"hh:mm:ss am",</c> + CRLF.
/// Failures are logged, not raised: the desktop's file write never stopped a login it had already accepted.
/// </summary>
public sealed class LoginLogFile(IOptions<SessionTrackingOptions> options, ILogger<LoginLogFile> log) : ILoginLog
{
    private static readonly SemaphoreSlim Gate = new(1, 1);

    /// <summary>The line written for a login.</summary>
    public static string Line(string userCode, DateTime now) =>
        $"\"Login\",\"{userCode}\",#{now:yyyy-MM-dd}#,\"{now.ToString("hh:mm:ss tt", CultureInfo.InvariantCulture).ToLowerInvariant()}\",\r\n";

    /// <inheritdoc />
    public async Task AppendAsync(string userCode, DateTime now)
    {
        var path = options.Value.LoginLogPath;
        if (string.IsNullOrWhiteSpace(path)) return;
        await Gate.WaitAsync();
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            await File.AppendAllTextAsync(path, Line(userCode, now));
        }
        catch (Exception e) when (e is IOException or UnauthorizedAccessException)
        {
            log.LogWarning(e, "Could not append to the login log {Path}", path);
        }
        finally
        {
            Gate.Release();
        }
    }
}
