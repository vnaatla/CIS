using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Payment;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Payment > To Member (frmPaymentMemberAuto08, frmPrintPVRange / frmPrintChequeRange in member mode).</summary>
[Route("api/payment/member")]
[RequireAccess(AccessRight.PaymentMember)]
public sealed class PayMbmController(IPayMbmService pay, IPayClinicService clinic) : ApiControllerBase
{
    /// <summary>Banks and group companies.</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups()
    {
        var (banks, companies, _) = await clinic.LookupsAsync();
        return new { banks = banks.Select(b => new { code = b.Code, name = b.Name }), companies };
    }

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(PayMbmSearchRequest request) => pay.SearchAsync(request);

    /// <summary>The PV just saved (pvNo) or the case just updated (caseNo).</summary>
    [HttpGet("relist")]
    public Task<List<Dictionary<string, object?>>> Relist([FromQuery] string? pvNo, [FromQuery] string? caseNo) =>
        pay.SearchAsync(new PayMbmSearchRequest { Bordx = null }, pvNo, caseNo);

    /// <summary>Save: one PV for the ticked cases.</summary>
    [HttpPost("save")]
    public async Task<object> Save(PayMbmSaveRequest request) => new { pvNo = await pay.SaveAsync(request, UserCode), message = PayMbmService.Saved };

    /// <summary>Update the selected PV.</summary>
    [HttpPost("update")]
    public async Task<object> Update(PayMbmSaveRequest request)
    {
        var (pvNo, caseNo) = await pay.UpdateAsync(request);
        return new { pvNo, caseNo, message = PayMbmService.Updated };
    }

    /// <summary>Preview PV (ticked cases) or the PV range print.</summary>
    [HttpPost("voucher")]
    public async Task<FileContentResult> Voucher(PayMbmVoucherRequest r) =>
        Download(await pay.VoucherAsync(r.Cases, r.PvFrom, r.PvTo, r.PayTo, r.PvNo, UserCode));

    /// <summary>The member PVs of a range.</summary>
    [HttpPost("pv-range")]
    public Task<List<Dictionary<string, object?>>> PvRange(PayClnInvRangeRequest request) => pay.PvRangeAsync(request);

    /// <summary>The member cheques of a range (account: Print Acc No).</summary>
    [HttpPost("cheques")]
    public Task<List<Dictionary<string, object?>>> Cheques(PayClnInvRangeRequest request, [FromQuery] bool account = false) => pay.ChequesAsync(request, account);

    /// <summary>Print the member cheques (marked printed unless account).</summary>
    [HttpPost("cheques/print")]
    public async Task<FileContentResult> PrintCheques(PayClnInvRangeRequest request, [FromQuery] bool account = false) =>
        Download(await pay.PrintChequesAsync(request, account));

    private FileContentResult Download(Application.Documents.FileOutput f) => File(f.Content, f.ContentType, f.FileName);
}
