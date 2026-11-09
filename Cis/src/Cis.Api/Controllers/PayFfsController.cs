using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Payment;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Payment > FFS Payment (FrmPaymentFFS).</summary>
[Route("api/payment/ffs")]
[RequireAccess(AccessRight.PaymentFfs)]
public sealed class PayFfsController(IPayFfsService ffs, IPayClinicService clinic) : ApiControllerBase
{
    /// <summary>Banks and group companies (payees).</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups()
    {
        var (banks, companies, _) = await clinic.LookupsAsync();
        return new { banks = banks.Select(b => new { code = b.Code, name = b.Name }), companies };
    }

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<PayFfsResult> Search(PayFfsSearchRequest request) => ffs.SearchAsync(request);

    /// <summary>The PV just saved / updated.</summary>
    [HttpGet("pv/{pvNo}")]
    public Task<PayFfsResult> Pv(string pvNo) => ffs.SearchAsync(new PayFfsSearchRequest { Scan = false }, pvNo);

    /// <summary>Save: a new PV for the ticked bordereaux.</summary>
    [HttpPost("save")]
    public async Task<object> Save(PayFfsSaveRequest request) => new { pvNo = await ffs.SaveAsync(request, UserCode), message = PayFfsService.Saved };

    /// <summary>Update the selected PV row.</summary>
    [HttpPost("update")]
    public async Task<object> Update(PayFfsSaveRequest request)
    {
        await ffs.UpdateAsync(request, UserCode);
        return new { message = PayFfsService.Saved };
    }

    /// <summary>Print PV (pvNo) / Preview PV (ticked bordereaux).</summary>
    [HttpPost("voucher")]
    public async Task<FileContentResult> Voucher(PayFfsSaveRequest request)
    {
        var f = await ffs.VoucherAsync(request.PvNo, request.Ticked, UserCode);
        return File(f.Content, f.ContentType, f.FileName);
    }
}
