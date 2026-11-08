using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Mco;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>MCO Fees > Fee Listing (frmMCOFeeListNew).</summary>
[Route("api/mco-fees/fee-listing")]
[RequireAccess(AccessRight.McoFeeListing)]
public sealed class McoListController(IMcoListService mco) : ApiControllerBase
{
    /// <summary>Payors and group companies.</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups()
    {
        var (payors, groups) = await mco.LookupsAsync();
        return new { payors = payors.Select(p => new { code = p.Code, name = p.Name }), groups = groups.Select(g => new { code = g.Code, name = g.Name }) };
    }

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<McoListResult> Search(McoListRequest request) => mco.SearchAsync(request);

    /// <summary>Print / Print Preview (PDF). Print numbers and stamps a new MCO invoice.</summary>
    [HttpPost("print")]
    public async Task<FileContentResult> Print(McoListPrintRequest request)
    {
        var file = await mco.PrintAsync(request, UserCode);
        return File(file.Content, file.ContentType, file.FileName);
    }
}
