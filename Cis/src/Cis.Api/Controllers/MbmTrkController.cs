using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Membership > MBM Tracking (FrmNewMBMTracking).</summary>
[Route("api/tracking/member")]
[RequireAccess(AccessRight.MbmTracking)]
public sealed class MbmTrkController(IMbmTrkService tracking) : ApiControllerBase
{
    /// <summary>The report catalogue.</summary>
    [HttpGet("reports")]
    public IReadOnlyList<MbmTrkReport> Reports() => tracking.Reports();

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<MbmTrkResult> Search(MbmTrkRequest request) => tracking.SearchAsync(request);
}
