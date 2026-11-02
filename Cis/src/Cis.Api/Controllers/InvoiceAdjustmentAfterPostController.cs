using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Clinic > Adjustment (After Posting) (frmInvoiceAdjAfterPost): a sub-case for a case posted to SUN.</summary>
[Route("api/invoice/adjustment-after-posting")]
[RequireAccess(AccessRight.InvoiceAdjustment)]
public sealed class InvoiceAdjustmentAfterPostController(IInvoiceAdjustmentService adjustment) : ApiControllerBase
{
    /// <summary>Go.</summary>
    [HttpGet("{casNumber}")]
    public Task<AdjustmentCase> Load(string casNumber) => adjustment.LoadAfterPostAsync(casNumber);

    /// <summary>Compute: the worksheet of the changed case.</summary>
    [HttpPost("compute")]
    public Task<WorksheetResult> Compute(AdjustmentRequest request) => adjustment.ComputeAsync(request, UserCode);

    /// <summary>Create: the sub-case.</summary>
    [HttpPost]
    public Task<AdjustmentResult> Create(AdjustmentRequest request) => adjustment.CreateAfterPostAsync(request, UserCode);
}
