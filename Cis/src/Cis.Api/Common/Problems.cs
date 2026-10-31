using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Common;

/// <summary>Builds the project's Problem Details shape: RFC 7807 fields plus <c>message</c> (shown by the UI) and <c>traceId</c>.</summary>
public static class Problems
{
    /// <summary>A Problem Details body with the user-facing <paramref name="message"/>.</summary>
    public static ProblemDetails Create(HttpContext context, int status, string title, string message) => new()
    {
        Status = status,
        Title = title,
        Detail = message,
        Extensions = { ["message"] = message, ["traceId"] = context.TraceIdentifier },
    };

    /// <summary>400 response for a request that failed its validator.</summary>
    public static ObjectResult Invalid(this ControllerBase controller, string message) =>
        new(Create(controller.HttpContext, StatusCodes.Status400BadRequest, "Request validation failed", message))
        {
            StatusCode = StatusCodes.Status400BadRequest,
            ContentTypes = { "application/problem+json" },
        };
}
