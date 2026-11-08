using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Mco;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>MCO Fees > MCO Invoice GST Zeroised (frmMCOFeeGSTZeroised).</summary>
[Route("api/mco-fees/invoice-gst-zeroised")]
[RequireAccess(AccessRight.McoInvoice)]
public sealed class McoZeroController(IMcoZeroService zero) : ApiControllerBase
{
    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<McoGstResult> Search(McoZeroRequest request) => zero.SearchAsync(request);

    /// <summary>Print: number and stamp the checked members, then the zeroised invoice (Excel).</summary>
    [HttpPost("print")]
    public async Task<FileContentResult> Print(McoZeroRequest request)
    {
        var file = await zero.PrintAsync(request, UserCode);
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Print Excel: the zeroised invoice of the searched invoice no.</summary>
    [HttpPost("print-invoice")]
    public async Task<FileContentResult> PrintInvoice(McoZeroRequest request)
    {
        var file = await zero.PrintInvoiceAsync(request);
        return File(file.Content, file.ContentType, file.FileName);
    }
}
