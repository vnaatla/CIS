using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim / Bordx > Claim Tracking 2 (FrmClaimRpt2).</summary>
[Route("api/tracking/claim2")]
[RequireAccess(AccessRight.ClaimTracking2)]
public sealed class ClmTrk2Controller(IClmTrk2Service tracking, IClmTrkService lookups) : ClmTrk2ControllerBase(tracking, lookups, "main");

/// <summary>Claim / Bordx > Claim Tracking 2 (Jupiter) (FrmNewClaimRpt2); the desktop read the Jupiter server.</summary>
[Route("api/tracking/claim2-jupiter")]
[RequireAccess(AccessRight.ClaimTracking2Jupiter)]
public sealed class ClmTrk2JupiterController(IClmTrk2Service tracking, IClmTrkService lookups) : ClmTrk2ControllerBase(tracking, lookups, "jupiter");

/// <summary>The Claim Tracking 2 endpoints of one form.</summary>
public abstract class ClmTrk2ControllerBase(IClmTrk2Service tracking, IClmTrkService lookups, string variant) : ApiControllerBase
{
    /// <summary>A combo: the same lists as Claim Tracking (SearchCLN, SearchCLNGrp, SearchPolStatus, SearchState, payors, group companies).</summary>
    [HttpGet("lookup/{kind}")]
    public Task<List<string>> Lookup(string kind, [FromQuery] string? key) => lookups.LookupAsync(kind, key);

    /// <summary>The report choices.</summary>
    [HttpGet("reports")]
    public IReadOnlyList<ClmTrk2Report> Reports() => tracking.Reports(variant);

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<ClmTrkResult> Search(ClmTrk2Request request) => tracking.SearchAsync(request, variant);
}
