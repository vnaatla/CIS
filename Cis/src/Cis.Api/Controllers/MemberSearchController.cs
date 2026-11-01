using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Membership;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Member search pop-ups (frmSearchMBM, frmSearchMBMIC, frmSearchMBM05) opened from MBM Adjustment / Enquiry / Topup and the invoices.</summary>
[Route("api/membership/search")]
[RequireAccess(AccessRight.MbmAdjustment, AccessRight.InvoiceClinic, AccessRight.InvoiceMember)]
public sealed class MemberSearchController(IMemberSearchService search) : ApiControllerBase
{
    /// <summary>frmSearchMBM: by Mem Num, Policy Num, Mem Name, IC Num, Employee Num or Policy Sub Num (at least 3 characters); IcExact = frmSearchMBMIC.</summary>
    [HttpGet]
    public Task<MemberSearchResult> Search([FromQuery] MemberSearchBy by, [FromQuery] string? text) => search.SearchAsync(by, text);

    /// <summary>frmSearchMBM05: member numbers starting with the caller's member number.</summary>
    [HttpGet("prefix")]
    public Task<MemberSearchResult> ByPrefix([FromQuery] string? memberNo) => search.ByPrefixAsync(memberNo);
}
