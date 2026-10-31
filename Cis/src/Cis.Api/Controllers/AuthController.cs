using Cis.Api.Auth;
using Cis.Api.Auth.Login;
using Cis.Api.Common;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Login, change password, logout (frmLogin, Exit > Logout / Close Connection).</summary>
[Route("api/auth")]
public sealed class AuthController(ILoginService login) : ApiControllerBase
{
    /// <summary>Signs in with a USR account and returns a JWT.</summary>
    [AllowAnonymous]
    [HttpPost("login")]
    public Task<SessionResult> Login(LoginRequest request) => login.LoginAsync(request, RemoteIp);

    /// <summary>Changes the password; answers the desktop's "Password Update Sucessfully".</summary>
    [AllowAnonymous]
    [HttpPost("change-password")]
    public async Task<object> ChangePassword(ChangePasswordRequest request) => new { message = await login.ChangePasswordAsync(request, RemoteIp) };

    /// <summary>Ends the session (USR.CIScnt - 1).</summary>
    [HttpPost("logout")]
    public async Task<IActionResult> Logout()
    {
        await login.LogoutAsync(UserCode);
        return NoContent();
    }

    /// <summary>The signed-in user (from the token).</summary>
    [HttpGet("me")]
    public SessionUser Me() => new(UserCode, User.FindFirst("name")?.Value ?? "", User.FindFirst(TokenService.AccessClaim)?.Value ?? "");
}
