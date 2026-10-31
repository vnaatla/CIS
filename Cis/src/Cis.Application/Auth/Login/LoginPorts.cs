using Cis.Application.Abstractions;
using Cis.Domain.Users;

namespace Cis.Application.Auth.Login;

/// <summary>Data access of login / logout / change password (HISMaintenance.dbo.USR and Tbl_USRTrack).</summary>
public interface ILoginRepository
{
    /// <summary>The active (USRStatus = 'A') user, or null.</summary>
    Task<UserAccount?> GetActiveUserAsync(IDbSession session, string userCode);

    /// <summary><c>UPDATE USR SET USRStatus = 'S'</c> (third wrong password).</summary>
    Task SuspendAsync(IDbSession session, string userCode);

    /// <summary>CIScnt + 1 and a Tbl_USRTrack row (SYSName 'CIS') with the desktop's text time and date.</summary>
    Task RecordLoginAsync(IDbSession session, string userCode, string loginTime, string loginDate);

    /// <summary>CIScnt - 1 (Logout / Close Connection).</summary>
    Task RecordLogoutAsync(IDbSession session, string userCode);

    /// <summary>Stores the scrambled new password.</summary>
    Task ChangePasswordAsync(IDbSession session, string userCode, string scrambledPassword);
}

/// <summary>
/// Wrong-password counter of frmLogin (<c>counter</c>, <c>txtCheckLogon</c>). The desktop kept it in the open login window; the web keeps it
/// per user code and client address for a while, which is the closest stateless equivalent.
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

/// <summary>The text login log (frmLogin.OutputToTxtfile, <c>DocPath\FileLogin\LoginCIS.txt</c>).</summary>
public interface ILoginLog
{
    /// <summary>Appends a "Login" line for the user.</summary>
    Task AppendAsync(string userCode, DateTime now);
}

/// <summary>Issues the session token (JWT) carrying the user's code, name and USRAccess flags.</summary>
public interface ITokenIssuer
{
    /// <summary>A signed token for the user and its expiry (UTC).</summary>
    (string Token, DateTime ExpiresUtc) Create(string userCode, string userName, string access);
}
