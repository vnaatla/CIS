using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Payment;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Payment > To Clinic (frmPaymentClinic).</summary>
[Route("api/payment/clinic")]
[RequireAccess(AccessRight.PaymentClinic)]
public sealed class PayClinicController(IPayClinicService pay) : ApiControllerBase
{
    /// <summary>Banks (code, name, PV code), group companies, clinics.</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups()
    {
        var (banks, companies, clinics) = await pay.LookupsAsync();
        return new
        {
            banks = banks.Select(b => new { code = b.Code, name = b.Name, pvCode = b.PvCode }), companies,
            clinics = clinics.Select(c => new { code = c.Code, name = c.Name }),
        };
    }

    /// <summary>Search (or the upload's invoice numbers).</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(PayClinicSearchRequest request) => pay.SearchAsync(request);

    /// <summary>The rows of one PV.</summary>
    [HttpGet("pv/{pvNo}")]
    public Task<List<Dictionary<string, object?>>> Pv(string pvNo) => pay.SearchAsync(new PayClinicSearchRequest(), pvNo);

    /// <summary>Add: a new PV for the ticked cases.</summary>
    [HttpPost("add")]
    public async Task<object> Add(PayClinicSaveRequest request)
    {
        var pvNo = await pay.AddAsync(request, UserCode);
        return new { pvNo, message = PayClinicService.Saved };
    }

    /// <summary>Update the edited PV.</summary>
    [HttpPost("update")]
    public async Task<object> Update(PayClinicSaveRequest request)
    {
        await pay.UpdateAsync(request, UserCode);
        return new { message = PayClinicService.Saved };
    }

    /// <summary>Print PV (pvNo) or Preview PV (ticked cases).</summary>
    [HttpPost("voucher")]
    public async Task<FileContentResult> Voucher(PayClinicVoucherRequest request)
    {
        var file = await pay.VoucherAsync(request.PvNo, request.Cases, UserCode);
        return File(file.Content, file.ContentType, file.FileName);
    }
}
