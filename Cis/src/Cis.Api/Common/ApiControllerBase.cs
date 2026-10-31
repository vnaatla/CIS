using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Common;

/// <summary>Base of the API controllers: the logged-in user's code (the JWT name claim) for audit columns and logs.</summary>
[ApiController]
public abstract class ApiControllerBase : ControllerBase
{
    /// <summary>USR.USRCode of the caller.</summary>
    protected string UserCode => User.Identity?.Name ?? "";

    /// <summary>USR.USRAccess of the caller (from the token), for services that check a right per record type.</summary>
    protected string? UserAccess => User.FindFirst(Auth.TokenService.AccessClaim)?.Value;

    /// <summary>Remote address of the caller, for logs.</summary>
    protected string? RemoteIp => HttpContext.Connection.RemoteIpAddress?.ToString();
}
