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

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<ClmTrkResult> Search(ClmTrkRequest request) => tracking.SearchAsync(request);
}
