using Cis.Domain.Security;

namespace Cis.Domain.Users;

/// <summary>An active user of HISMaintenance.dbo.USR (the account the desktop's frmLogin checks).</summary>
/// <param name="UsrCode">USRCode.</param>
/// <param name="UsrName">USRName.</param>
/// <param name="UsrPassword">USRPassword (scrambled with <see cref="LegacyCrypt"/>).</param>
/// <param name="UsrAccess">USRAccess (position n = CheckAccess(n)).</param>
public record UserAccount(string UsrCode, string? UsrName, string? UsrPassword, string? UsrAccess)
{
    /// <summary>The desktop's password check: <c>UCase(Crypt(typed)) = UCase(USRPassword)</c>.</summary>
    public bool PasswordMatches(string? typed) => LegacyCrypt.Matches(typed ?? "", UsrPassword);

    /// <summary>VB6 CheckAccess(n) for this user.</summary>
    public bool HasAccess(int position) => AccessRight.Has(UsrAccess, position);
}
