using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim / Bordx > Claim Tracking (frmClaimRpt).</summary>
[Route("api/tracking/claim")]
[RequireAccess(AccessRight.ClaimTracking)]
public sealed class ClmTrkController(IClmTrkService tracking) : ApiControllerBase
{
    /// <summary>The report catalogue.</summary>
    [HttpGet("reports")]
    public IReadOnlyList<MbmTrkReport> Reports() => tracking.Reports();

    /// <summary>A combo (see IClmTrkService.LookupAsync).</summary>
    [HttpGet("lookup/{kind}")]
    public Task<List<string>> Lookup(string kind, [FromQuery] string? key) => tracking.LookupAsync(kind, key);

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<ClmTrkResult> Search(ClmTrkRequest request) => tracking.SearchAsync(request);
}
