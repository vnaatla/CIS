using Cis.Api.Common;
using FluentValidation;

namespace Cis.Api.Auth.Login;

/// <summary>cmdOK_Click: the user name is required (the password may be empty; it then simply does not match).</summary>
public sealed class LoginRequestValidator : RequestValidator<LoginRequest>
{
    /// <summary>Defines the rules.</summary>
    public LoginRequestValidator()
    {
        RuleFor(r => r.UserCode).Must(c => !string.IsNullOrWhiteSpace(c)).WithMessage(LoginMessages.EnterUserName);
    }
}

/// <summary>cmdSave_Click checks, in the desktop's order.</summary>
public sealed class ChangePasswordRequestValidator : RequestValidator<ChangePasswordRequest>
{
    /// <summary>Length of USR.USRPassword.</summary>
    public const int PasswordLength = 10;

    /// <summary>Defines the rules.</summary>
    public ChangePasswordRequestValidator()
    {
        RuleFor(r => r.UserCode).Must(c => !string.IsNullOrWhiteSpace(c)).WithMessage(LoginMessages.EnterUserName);
        RuleFor(r => r.Password).Must(p => !string.IsNullOrEmpty(p)).WithMessage(LoginMessages.EnterPassword);
        RuleFor(r => r.NewPassword).Must(p => !string.IsNullOrEmpty(p)).WithMessage(LoginMessages.EnterNewPassword);
        RuleFor(r => r.ConfirmPassword).Must(p => !string.IsNullOrEmpty(p)).WithMessage(LoginMessages.EnterConfirmPassword);
        RuleFor(r => r).Must(r => r.NewPassword == r.ConfirmPassword).WithMessage(LoginMessages.PasswordsDiffer);
        RuleFor(r => r.NewPassword).MaximumLength(PasswordLength).WithMessage(LoginMessages.NewPasswordTooLong);
    }
}
