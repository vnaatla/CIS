using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Tracking: the shared actions.</summary>
public abstract class InvTrkControllerBase(IInvTrkService tracking) : ApiControllerBase
{
    /// <summary>A combo: clinics, payors, companies (key = payor), subsidiaries / departments (key = group company).</summary>
    [HttpGet("lookup/{kind}")]
    public Task<List<string>> Lookup(string kind, [FromQuery] string? key) => tracking.LookupAsync(kind, key);

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<InvTrkResult> Search(InvTrkRequest request) => tracking.SearchAsync(request);
}

/// <summary>Invoice > Tracking (FrmInvoiceTracking).</summary>
[Route("api/tracking/invoice")]
[RequireAccess(AccessRight.InvoiceTracking)]
public sealed class InvTrkController(IInvTrkService tracking) : InvTrkControllerBase(tracking);

