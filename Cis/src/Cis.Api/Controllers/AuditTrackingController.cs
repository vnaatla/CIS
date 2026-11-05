using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Audit;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>LG &amp; Audit > Audit Tracking (FrmAuditInvoiceTracking).</summary>
[Route("api/lg-audit/audit-tracking")]
[RequireAccess(AccessRight.AuditReview)]
public sealed class AuditTrackingController(IAuditTrackingService tracking) : ApiControllerBase
{
    /// <summary>Search: one row per medication of each matching case.</summary>
    [HttpPost("search")]
    public Task<List<AuditTrackingRow>> Search(AuditTrackingRequest filter) => tracking.SearchAsync(filter);

    /// <summary>Print to File: the audit tracking Excel.</summary>
    [HttpPost("print")]
    public async Task<FileContentResult> Print(AuditTrackingRequest filter)
    {
        var file = await tracking.PrintAsync(filter);
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Subsidiaries and departments of a group company.</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups([FromQuery] string? groupCompany)
    {
        var (subsidiaries, departments) = await tracking.LookupsAsync(groupCompany);
        return new { subsidiaries, departments };
    }
}
