using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim / Bordx > Claim Tracking (frmClaimRpt).</summary>
[Route("api/tracking/claim")]
[RequireAccess(AccessRight.ClaimTracking)]
public sealed class ClmTrkController(IClmTrkService tracking) : ClmTrkControllerBase(tracking, "main");

/// <summary>Claim Tracking (Simple) (frmClaimRptSimple): CISClaimTracking_Simple.</summary>
[Route("api/tracking/claim-simple")]
[RequireAccess(AccessRight.ClaimTracking)]
public sealed class ClmTrkSimpleController(IClmTrkService tracking) : ClmTrkControllerBase(tracking, "simple");

/// <summary>Claim Tracking (Jupiter) (frmNewClaimRpt); the desktop read the Jupiter server.</summary>
[Route("api/tracking/claim-jupiter")]
[RequireAccess(AccessRight.ClaimTrackingJupiter)]
public sealed class ClmTrkJupiterController(IClmTrkService tracking) : ClmTrkControllerBase(tracking, "jupiter");

/// <summary>Claim Tracking (Jupiter - Simple) (frmNewClaimRptSimpleJup).</summary>
[Route("api/tracking/claim-jupiter-simple")]
[RequireAccess(AccessRight.ClaimTrackingJupiter)]
public sealed class ClmTrkJupiterSimpleController(IClmTrkService tracking) : ClmTrkControllerBase(tracking, "jupitersimple");

/// <summary>Claim Tracking (SUN System) (frmClaimRptSUN).</summary>
[Route("api/tracking/claim-sun")]
[RequireAccess(AccessRight.ClaimTracking)]
public sealed class ClmTrkSunController(IClmTrkService tracking) : ClmTrkControllerBase(tracking, "sun");

/// <summary>The claim tracking endpoints of one variant form.</summary>
public abstract class ClmTrkControllerBase(IClmTrkService tracking, string variant) : ApiControllerBase
{
    /// <summary>The report catalogue.</summary>
    [HttpGet("reports")]
    public IReadOnlyList<MbmTrkReport> Reports() => tracking.Reports(variant);

    /// <summary>A combo (see IClmTrkService.LookupAsync).</summary>
    [HttpGet("lookup/{kind}")]
    public Task<List<string>> Lookup(string kind, [FromQuery] string? key) => tracking.LookupAsync(kind, key);

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<ClmTrkResult> Search(ClmTrkRequest request) => tracking.SearchAsync(request, variant);
}
