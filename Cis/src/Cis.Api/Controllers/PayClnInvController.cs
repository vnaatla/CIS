using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Payment;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Payment > To Clinic - Invoice (frmPaymentCLNInvoice, frmInvoiceDetPopup, frmPrintPVRange, frmPrintChequeRange).</summary>
[Route("api/payment/clinic-invoice")]
[RequireAccess(AccessRight.PaymentClinicInvoice)]
public sealed class PayClnInvController(IPayClnInvService pay) : ApiControllerBase
{
    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(PayClnInvSearchRequest request) => pay.SearchAsync(request);

    /// <summary>The PVs just saved / updated.</summary>
    [HttpGet("pv")]
    public Task<List<Dictionary<string, object?>>> Pv([FromQuery] string pvNo, [FromQuery] string? pvLast) => pay.SearchAsync(new PayClnInvSearchRequest(), pvNo, pvLast);

    /// <summary>The invoices of a row (frmInvoiceDetPopup).</summary>
    [HttpGet("invoices")]
    public Task<List<Dictionary<string, object?>>> Invoices([FromQuery] string clinic, [FromQuery] string month, [FromQuery] string year) =>
        pay.InvoicesAsync(clinic, month, year);

    /// <summary>Save: one PV per clinic of the ticked rows.</summary>
    [HttpPost("save")]
    public async Task<object> Save(PayClnInvSaveRequest request)
    {
        var (first, last) = await pay.SaveAsync(request, UserCode);
        return new { first, last, message = PayClnInvService.Saved };
    }

    /// <summary>Update the edited PV.</summary>
    [HttpPost("update")]
    public async Task<object> Update(PayClnInvSaveRequest request)
    {
        await pay.UpdateAsync(request, UserCode);
        return new { message = PayClnInvService.Saved };
    }

    /// <summary>Print PV (pvNo) / Preview PV (ticked rows).</summary>
    [HttpPost("voucher")]
    public async Task<FileContentResult> Voucher(PayClnInvSaveRequest request) =>
        Download(await pay.VoucherAsync(request.PvNo, request.Ticked, UserCode));

    /// <summary>frmPrintPVRange: the PVs of a range.</summary>
    [HttpPost("pv-range")]
    public Task<List<Dictionary<string, object?>>> PvRange(PayClnInvRangeRequest request) => pay.PvRangeAsync(request);

    /// <summary>frmPrintPVRange: print the range.</summary>
    [HttpPost("pv-range/print")]
    public async Task<FileContentResult> PrintPvRange(PayClnInvRangeRequest request) => Download(await pay.PrintPvRangeAsync(request, UserCode));

    /// <summary>frmPrintChequeRange: the cheques of a range.</summary>
    [HttpPost("cheques")]
    public Task<List<Dictionary<string, object?>>> Cheques(PayClnInvRangeRequest request) => pay.ChequesAsync(request);

    /// <summary>Print the cheques (marks them printed).</summary>
    [HttpPost("cheques/print")]
    public async Task<FileContentResult> PrintCheques(PayClnInvRangeRequest request) => Download(await pay.PrintChequesAsync(request));

    private FileContentResult Download(Application.Documents.FileOutput f) => File(f.Content, f.ContentType, f.FileName);
}
