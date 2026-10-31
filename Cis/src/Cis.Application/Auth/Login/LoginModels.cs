namespace Cis.Application.Auth.Login;

/// <summary>Login box (frmLogin: txtUserName, txtPassword).</summary>
/// <param name="UserCode">USR.USRCode.</param>
/// <param name="Password">Plain password as typed.</param>
public record LoginRequest(string UserCode, string Password);

/// <summary>Change-password panel of frmLogin (txtUserName, txtPassword, txtnewpass, txtnewconfirm).</summary>
/// <param name="UserCode">USR.USRCode.</param>
/// <param name="Password">Current password.</param>
/// <param name="NewPassword">New password.</param>
/// <param name="ConfirmPassword">New password typed again.</param>
public record ChangePasswordRequest(string UserCode, string Password, string NewPassword, string ConfirmPassword);

/// <summary>The signed-in user.</summary>
/// <param name="UserCode">USR.USRCode, upper case.</param>
/// <param name="UserName">USR.USRName.</param>
/// <param name="Access">USR.USRAccess flags (position n = CheckAccess(n)).</param>
public record SessionUser(string UserCode, string UserName, string Access);

/// <summary>Answer of a successful login.</summary>
/// <param name="Token">JWT to send as <c>Authorization: Bearer</c>.</param>
/// <param name="ExpiresUtc">Token expiry.</param>
/// <param name="User">The signed-in user.</param>
public record SessionResult(string Token, DateTime ExpiresUtc, SessionUser User);

/// <summary>Messages of frmLogin, word for word.</summary>
public static class LoginMessages
{
    /// <summary>cmdOK_Click / cmdSave_Click with an empty user name.</summary>
    public const string EnterUserName = "Please Enter User Name";
    /// <summary>cmdSave_Click with an empty password.</summary>
    public const string EnterPassword = "Please Enter Password";
    /// <summary>Unknown or inactive user.</summary>
    public const string InvalidUserName = "Invalid User Name!! Please enter again..";
    /// <summary>Wrong password.</summary>
    public const string InvalidPassword = "Invalid Password!! Please enter again..";
    /// <summary>Fourth attempt after three wrong passwords for the same user.</summary>
    public const string Suspended = "Your Account Been Suspended";
    /// <summary>cmdSave_Click with an empty new password.</summary>
    public const string EnterNewPassword = "Please Enter Valid New Password";
    /// <summary>cmdSave_Click with an empty confirmation.</summary>
    public const string EnterConfirmPassword = "Please Enter Valid Re-Type Password";
    /// <summary>cmdSave_Click when the two new passwords differ.</summary>
    public const string PasswordsDiffer = "Password Verification Is In Valid";
    /// <summary>cmdSave_Click when the user / current password do not match.</summary>
    public const string NoMatch = "No Record Match For User ID or Password";
    /// <summary>cmdSave_Click success.</summary>
    public const string PasswordUpdated = "Password Update Sucessfully";
    /// <summary>Web only: USR.USRPassword holds 10 characters.</summary>
    public const string NewPasswordTooLong = "New password must be at most 10 characters.";
}
