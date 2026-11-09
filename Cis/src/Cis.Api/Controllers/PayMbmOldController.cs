using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Payment;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Payment > To Member (old) (frmPaymentMember). The preview, cheque and range prints are those of api/payment/member.</summary>
[Route("api/payment/member-legacy")]
[RequireAccess(AccessRight.PaymentMember)]
public sealed class PayMbmOldController(IPayMbmOldService pay) : ApiControllerBase
{
    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(PayMbmSearchRequest request) => pay.SearchAsync(request);

    /// <summary>The PV just saved / updated.</summary>
    [HttpGet("pv/{pvNo}")]
    public Task<List<Dictionary<string, object?>>> Pv(string pvNo) => pay.SearchAsync(new PayMbmSearchRequest { Scan = false, Bordx = null }, pvNo);

    /// <summary>Save: one PV for the ticked cases with the frame's details.</summary>
    [HttpPost("save")]
    public async Task<object> Save(PayMbmOldSaveRequest request) => new { pvNo = await pay.SaveAsync(request, UserCode), message = PayMbmOldService.Saved };

    /// <summary>Update the selected row.</summary>
    [HttpPost("update")]
    public async Task<object> Update(PayMbmOldSaveRequest request) => new { message = await pay.UpdateAsync(request, UserCode) };

    /// <summary>Print PV.</summary>
    [HttpGet("voucher")]
    public async Task<FileContentResult> Voucher([FromQuery] string? pvNo, [FromQuery] string? payTo)
    {
        var f = await pay.VoucherAsync(pvNo, payTo, UserCode);
        return File(f.Content, f.ContentType, f.FileName);
    }
}
