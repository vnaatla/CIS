using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim / Bordx > Claim Tracking 2 (FrmClaimRpt2).</summary>
[Route("api/tracking/claim2")]
[RequireAccess(AccessRight.ClaimTracking2)]
public sealed class ClmTrk2Controller(IClmTrk2Service tracking) : ClmTrk2ControllerBase(tracking, "main");

/// <summary>Claim / Bordx > Claim Tracking 2 (Jupiter) (FrmNewClaimRpt2); the desktop read the Jupiter server.</summary>
[Route("api/tracking/claim2-jupiter")]
[RequireAccess(AccessRight.ClaimTracking2Jupiter)]
public sealed class ClmTrk2JupiterController(IClmTrk2Service tracking) : ClmTrk2ControllerBase(tracking, "jupiter");

/// <summary>The Claim Tracking 2 endpoints of one form; the combos are those of Claim Tracking (api/tracking/claim/lookup).</summary>
public abstract class ClmTrk2ControllerBase(IClmTrk2Service tracking, string variant) : ApiControllerBase
{
    /// <summary>The report choices.</summary>
    [HttpGet("reports")]
    public IReadOnlyList<ClmTrk2Report> Reports() => tracking.Reports(variant);

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<ClmTrkResult> Search(ClmTrk2Request request) => tracking.SearchAsync(request, variant);
}
