using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Claim;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim > Pass to Claim (frmPass2Claim).</summary>
[Route("api/claim/pass-to-claim")]
[RequireAccess(AccessRight.PassToClaim)]
public sealed class PassToClaimController(IPassToClaimService pass) : ApiControllerBase
{
    /// <summary>Users who passed cases (Pass By combo).</summary>
    [HttpGet("users")]
    public Task<List<string>> Users() => pass.UsersAsync();

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<PassRow>> Search(PassFilters filters) => pass.SearchAsync(filters);

    /// <summary>Print: the list as Excel.</summary>
    [HttpPost("export")]
    public async Task<FileContentResult> Export(PassFilters filters)
    {
        var file = await pass.ExportAsync(filters);
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Pass the checked cases.</summary>
    [HttpPost]
    public async Task<object> Pass(PassRequest request)
    {
        var (message, rows) = await pass.PassAsync(request, UserCode);
        return new { message, rows };
    }
}
