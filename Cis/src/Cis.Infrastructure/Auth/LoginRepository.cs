using Cis.Application.Abstractions;
using Cis.Application.Auth.Login;
using Cis.Domain.Users;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Auth;

/// <summary><see cref="ILoginRepository"/> on the <c>cis.Login_*</c> procedures (HISMaintenance.dbo.USR / Tbl_USRTrack through synonyms).</summary>
public sealed class LoginRepository : ILoginRepository
{
    /// <inheritdoc />
    public Task<UserAccount?> GetActiveUserAsync(IDbSession session, string userCode) =>
        session.FirstOrDefaultAsync<UserAccount>("cis.Login_GetActiveUser", new { userCode });

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
