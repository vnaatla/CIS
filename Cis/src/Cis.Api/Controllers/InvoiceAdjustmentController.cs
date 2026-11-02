using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Clinic > Adjustment (frmInvoiceAdj, frmCasePending).</summary>
[Route("api/invoice/adjustment")]
[RequireAccess(AccessRight.InvoiceAdjustment)]
public sealed class InvoiceAdjustmentController(IInvoiceAdjustmentService adjustment) : ApiControllerBase
{
    /// <summary>Go: the case, what Edit allows, pending reasons, the member and the stored worksheet.</summary>
    [HttpGet("{casNumber}")]
    public Task<AdjustmentCase> Load(string casNumber) => adjustment.LoadAsync(casNumber);

    /// <summary>Compute: the worksheet of the changed case.</summary>
    [HttpPost("compute")]
    public Task<WorksheetResult> Compute(AdjustmentRequest request) => adjustment.ComputeAsync(request, UserCode);

    /// <summary>Update.</summary>
    [HttpPost]
    public Task<AdjustmentResult> Update(AdjustmentRequest request) => adjustment.UpdateAsync(request, UserCode);

    /// <summary>Audit review.</summary>
    [HttpPost("{casNumber}/audit")]
    public async Task<object> Audit(string casNumber, AuditReviewRequest request) => new { message = await adjustment.AuditAsync(casNumber, request) };

    /// <summary>Claims approval flag.</summary>
    [HttpPost("{casNumber}/claims-app")]
    public async Task<IActionResult> ClaimsApp(string casNumber)
    {
        await adjustment.ClaimsAppAsync(casNumber);
        return NoContent();
    }

    /// <summary>Void bordereaux (reimbursements).</summary>
    [HttpPost("{casNumber}/void-bordx")]
    public async Task<object> VoidBordx(string casNumber) => new { message = await adjustment.VoidBordxAsync(casNumber, UserCode) };
}
