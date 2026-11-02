using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Clinic > Adjustment (After Bordx) (frmInvoiceAdjAftBordx).</summary>
[Route("api/invoice/adjustment-after-bordx")]
[RequireAccess(AccessRight.InvoiceAdjustmentAfterBordx)]
public sealed class InvoiceAdjustmentAfterBordxController(IInvoiceAdjustmentService adjustment) : ApiControllerBase
{
    /// <summary>Go.</summary>
    [HttpGet("{casNumber}")]
    public Task<AdjustmentCase> Load(string casNumber) => adjustment.LoadAfterBordxAsync(casNumber);

    /// <summary>Compute: the worksheet of the changed case.</summary>
    [HttpPost("compute")]
    public Task<WorksheetResult> Compute(AdjustmentRequest request) => adjustment.ComputeAsync(request, UserCode);

    /// <summary>Update: the case and its invoice.</summary>
    [HttpPost]
    public Task<AdjustmentResult> Update(AdjustmentRequest request) => adjustment.UpdateAfterBordxAsync(request, UserCode);
}
