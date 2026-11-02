using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>The worksheet pop-up (frmWorksheet) of the invoice and adjustment screens.</summary>
[Route("api/invoice/worksheet")]
[RequireAccess(AccessRight.InvoiceClinic, AccessRight.InvoiceMember, AccessRight.InvoiceAdjustment, AccessRight.InvoiceAdjustmentAfterBordx)]
public sealed class WorksheetController(IWorksheetService worksheet) : ApiControllerBase
{
    /// <summary>The worksheet of an invoice row.</summary>
    [HttpPost("calculate")]
    public Task<WorksheetResult> Calculate(WorksheetInput input) => worksheet.CalculateAsync(input, UserCode);
}
