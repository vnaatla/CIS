using Cis.Api.Data;

namespace Cis.Api.Auth.Login;

/// <summary>Data access of login / logout / change password (HISMaintenance.dbo.USR and Tbl_USRTrack through synonyms).</summary>
public interface ILoginRepository
{
    /// <summary>The active (USRStatus = 'A') user, or null.</summary>
    Task<UserRecord?> GetActiveUserAsync(IDbSession session, string userCode);

    /// <summary><c>UPDATE USR SET USRStatus = 'S'</c> (third wrong password).</summary>
    Task SuspendAsync(IDbSession session, string userCode);

    /// <summary>CIScnt + 1 and a Tbl_USRTrack row (SYSName 'CIS') with the desktop's text time and date.</summary>
    Task RecordLoginAsync(IDbSession session, string userCode, string loginTime, string loginDate);

    /// <summary>CIScnt - 1 (Logout / Close Connection).</summary>
    Task RecordLogoutAsync(IDbSession session, string userCode);

    /// <summary>Stores the scrambled new password.</summary>
    Task ChangePasswordAsync(IDbSession session, string userCode, string scrambledPassword);
}

/// <inheritdoc />
public sealed class LoginRepository : ILoginRepository
{
    /// <inheritdoc />
    public Task<UserRecord?> GetActiveUserAsync(IDbSession session, string userCode) =>
        session.FirstOrDefaultAsync<UserRecord>("cis.Login_GetActiveUser", new { userCode });

    /// <inheritdoc />
    public Task SuspendAsync(IDbSession session, string userCode) =>
        session.ExecAsync("cis.Login_Suspend", new { userCode });

    /// <inheritdoc />
    public Task RecordLoginAsync(IDbSession session, string userCode, string loginTime, string loginDate) =>
        session.ExecAsync("cis.Login_Record", new { userCode, loginTime, loginDate });

    /// <inheritdoc />
    public Task RecordLogoutAsync(IDbSession session, string userCode) =>
        session.ExecAsync("cis.Login_Logout", new { userCode });

    /// <inheritdoc />
    public Task ChangePasswordAsync(IDbSession session, string userCode, string scrambledPassword) =>
        session.ExecAsync("cis.Login_ChangePassword", new { userCode, password = scrambledPassword });
}
