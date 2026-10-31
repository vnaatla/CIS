using System.Security.Claims;
using Cis.Api.Common;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

namespace Cis.Api.Auth;

/// <summary>Reads the menu rights carried by the JWT.</summary>
public static class ClaimsPrincipalAccess
{
    /// <summary>True when the user's token grants <paramref name="position"/> (<see cref="AccessRight"/>).</summary>
    public static bool HasAccess(this ClaimsPrincipal user, int position) =>
        AccessRight.Has(user.FindFirst(TokenService.AccessClaim)?.Value, position);
}

/// <summary>
/// Requires one of the given menu positions (several when screens share read-only endpoints). Class and method attributes stack:
/// both must pass. The refusal text is the desktop's "Access Denied!".
/// </summary>
[AttributeUsage(AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = true)]
public sealed class RequireAccessAttribute(int position, params int[] alternatives) : Attribute, IAuthorizationFilter
{
    /// <summary>The main position checked.</summary>
    public int Position => position;

    /// <summary>All positions that grant access.</summary>
    public IReadOnlyList<int> Positions => [position, .. alternatives];

    /// <inheritdoc />
    public void OnAuthorization(AuthorizationFilterContext context)
    {
        if (context.HttpContext.User.Identity?.IsAuthenticated != true) return;
        if (!Positions.Any(p => context.HttpContext.User.HasAccess(p)))
            context.Result = new ObjectResult(Problems.Create(context.HttpContext, StatusCodes.Status403Forbidden, "Forbidden", "Access Denied!"))
            {
                StatusCode = StatusCodes.Status403Forbidden,
            };
    }
}
