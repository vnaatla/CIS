using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Membership > MBM Tracking (FrmNewMBMTracking).</summary>
[Route("api/tracking/member")]
[RequireAccess(AccessRight.MbmTracking)]
public sealed class MbmTrkController(IMbmTrkService tracking) : MbmTrkControllerBase(tracking, "main");

/// <summary>Membership > MBM Tracking (Jupiter) (FrmNewMBMTracking2): its own report choices and columns; the desktop read the Jupiter server, here the same databases.</summary>
[Route("api/tracking/member-jupiter")]
[RequireAccess(AccessRight.MbmTrackingJupiter)]
public sealed class MbmTrkJupiterController(IMbmTrkService tracking) : MbmTrkControllerBase(tracking, "jupiter");

/// <summary>Membership > MBM Tracking (Simple) (frmNewMBMTrackingSimple): CISMBMTracking_Simple and Head Count.</summary>
[Route("api/tracking/member-simple")]
[RequireAccess(AccessRight.MbmTracking)]
public sealed class MbmTrkSimpleController(IMbmTrkService tracking) : MbmTrkControllerBase(tracking, "simple");

/// <summary>Membership > MBM Tracking (Jupiter - Simple) (frmNewMBMTrackingRepSimple).</summary>
[Route("api/tracking/member-jupiter-simple")]
[RequireAccess(AccessRight.MbmTrackingJupiter)]
public sealed class MbmTrkJupiterSimpleController(IMbmTrkService tracking) : MbmTrkControllerBase(tracking, "jupitersimple");

/// <summary>The member tracking endpoints of one variant form.</summary>
public abstract class MbmTrkControllerBase(IMbmTrkService tracking, string variant) : ApiControllerBase
{
    /// <summary>The report catalogue.</summary>
    [HttpGet("reports")]
    public IReadOnlyList<MbmTrkReport> Reports() => tracking.Reports(variant);

    /// <summary>A combo: states, ins, payors (key = health type), companies (key = payor), departments (key = group company).</summary>
    [HttpGet("lookup/{kind}")]
    public Task<List<string>> Lookup(string kind, [FromQuery] string? key) => tracking.LookupAsync(kind, key);

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<MbmTrkResult> Search(MbmTrkRequest request) => tracking.SearchAsync(request, variant);
}
