using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Common;

/// <summary>Base of the API controllers: the logged-in user's code (the JWT name claim) for audit columns and logs.</summary>
[ApiController]
public abstract class ApiControllerBase : ControllerBase
{
    /// <summary>USR.USRCode of the caller.</summary>
    protected string UserCode => User.Identity?.Name ?? "";

    /// <summary>Remote address of the caller, for logs.</summary>
    protected string? RemoteIp => HttpContext.Connection.RemoteIpAddress?.ToString();
}
