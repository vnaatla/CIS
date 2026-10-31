using System.Globalization;
using Cis.Api.Common;
using Cis.Api.Data;

namespace Cis.Api.Auth.Login;

/// <summary>Login, logout and change password with the desktop's rules (frmLogin, MDIMain.Logout_Click / mnuCloseConnection_Click).</summary>
public interface ILoginService
{
    /// <summary>Checks the user and password and opens a session (frmLogin.cmdOK_Click).</summary>
    Task<SessionResult> LoginAsync(LoginRequest request, string? client);

    /// <summary>Changes the password (frmLogin.cmdSave_Click) and answers the desktop's confirmation text.</summary>
    Task<string> ChangePasswordAsync(ChangePasswordRequest request, string? client);

    /// <summary>Ends the session: USR.CIScnt - 1 (Exit > Logout and Exit > Close Connection).</summary>
    Task LogoutAsync(string userCode);
}

/// <summary>
/// Desktop parity:
/// <list type="bullet">
/// <item>3 wrong passwords for the same user, then the next attempt sets USRStatus = 'S' whatever the password ("Your Account Been Suspended").</item>
/// <item>Success: CIScnt + 1, a Tbl_USRTrack row (SYSName 'CIS', LoginTime "h:mm:ss tt", LoginDate "dd/MM/yyyy" as in production), the LoginCIS.txt line.</item>
/// </list>
/// Deliberate differences (docs/screens/frmLogin.md): the plain password is <b>not</b> copied to USR.actpass (the desktop did, which leaks
/// passwords), and change password checks and stores the scrambled password (the desktop compared and stored plain text in HISDB, which
/// locked the user out).
/// </summary>
public sealed class LoginService(
    IDbSessions sessions,
    ILoginRepository repository,
    ILoginAttempts attempts,
    ILoginLog loginLog,
    TokenService tokens,
    TimeProvider clock,
    ILogger<LoginService> log) : ILoginService
{
    /// <summary>Wrong passwords allowed before the next attempt suspends the account (frmLogin: <c>counter = 3</c>).</summary>
    public const int MaxFailures = 3;

    /// <inheritdoc />
    public async Task<SessionResult> LoginAsync(LoginRequest request, string? client)
    {
        var typed = request.UserCode.Trim();
        if (attempts.Failures(typed, client) >= MaxFailures)
        {
            await using (var suspend = await sessions.BeginAsync())
            {
                await repository.SuspendAsync(suspend, typed);
                await suspend.CommitAsync();
            }
            attempts.Reset(typed, client);
            log.LogWarning("Account {UserCode} suspended after {Failures} wrong passwords from {Client}", Shorten(typed), MaxFailures, client);
            throw new UnauthorizedException(LoginMessages.Suspended);
        }

        await using var session = await sessions.BeginAsync();
        var user = await repository.GetActiveUserAsync(session, typed);
        if (user is null)
        {
            log.LogWarning("Login refused: unknown or inactive user {UserCode} from {Client}", Shorten(typed), client);
            throw new UnauthorizedException(LoginMessages.InvalidUserName);
        }
        if (!LegacyCrypt.Matches(request.Password ?? "", user.UsrPassword))
        {
            attempts.RegisterFailure(typed, client);
            log.LogWarning("Login refused: wrong password for {UserCode} from {Client}", user.UsrCode.Trim(), client);
            throw new UnauthorizedException(LoginMessages.InvalidPassword);
        }

        var now = clock.GetLocalNow().DateTime;
        await repository.RecordLoginAsync(session, typed,
            now.ToString("h:mm:ss tt", CultureInfo.InvariantCulture), now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture));
        await session.CommitAsync();
        attempts.Reset(typed, client);
        await loginLog.AppendAsync(request.UserCode, now);

        var code = user.UsrCode.Trim().ToUpperInvariant();
        var name = user.UsrName?.Trim() ?? "";
        var access = user.UsrAccess?.Trim() ?? "";
        var (token, expires) = tokens.Create(code, name, access);
        log.LogInformation("{UserCode} logged in from {Client}", code, client);
        return new SessionResult(token, expires, new SessionUser(code, name, access));
    }

    /// <inheritdoc />
    public async Task<string> ChangePasswordAsync(ChangePasswordRequest request, string? client)
    {
        await using var session = await sessions.BeginAsync();
        var user = await repository.GetActiveUserAsync(session, request.UserCode.Trim());
        if (user is null || !LegacyCrypt.Matches(request.Password, user.UsrPassword))
        {
            log.LogWarning("Password change refused for {UserCode} from {Client}", Shorten(request.UserCode), client);
            throw new UnauthorizedException(LoginMessages.NoMatch);
        }
        await repository.ChangePasswordAsync(session, user.UsrCode, LegacyCrypt.Encrypt(request.NewPassword).ToUpperInvariant());
        await session.CommitAsync();
        log.LogInformation("{UserCode} changed the password from {Client}", user.UsrCode.Trim(), client);
        return LoginMessages.PasswordUpdated;
    }

    /// <inheritdoc />
    public async Task LogoutAsync(string userCode)
    {
        await using var session = await sessions.BeginAsync();
        await repository.RecordLogoutAsync(session, userCode);
        await session.CommitAsync();
        log.LogInformation("{UserCode} logged out", userCode);
    }

    private static string Shorten(string? userCode) => (userCode ?? "").Trim() is { Length: > 30 } text ? text[..30] : (userCode ?? "").Trim();
}
