using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice pop-ups (frmSearchICD, frmCasList): open to the invoice screens' rights.</summary>
[Route("api/invoice/popups")]
[RequireAccess(AccessRight.InvoiceClinic, AccessRight.InvoiceMember, AccessRight.InvoiceAdjustment, AccessRight.InvoiceAdjustmentAfterBordx)]
public sealed class InvoicePopupController(IInvoicePopupService popups) : ApiControllerBase
{
    /// <summary>ICD / treatment reason search.</summary>
    [HttpGet("icd")]
    public Task<List<IcdRow>> Icd([FromQuery] string? code, [FromQuery] string? desc, [FromQuery] string? desc2, [FromQuery] string? criteria) =>
        popups.IcdAsync(code, desc, desc2, criteria);

    /// <summary>Case list: by 1 member no, 2 policy no, 3 name, 4 IC, 5 case no.</summary>
    [HttpGet("cases")]
    public Task<List<CaseListRow>> Cases([FromQuery] bool dental, [FromQuery] int by, [FromQuery] string? text) => popups.CasesAsync(dental, by, text);
}
