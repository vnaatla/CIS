using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Payment;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Payment > Tracking (FrmPayment) and Tracking (Jupiter) (frmPayment2, the same screen on the Jupiter connections).</summary>
[Route("api/payment/tracking")]
[RequireAccess(AccessRight.PaymentTracking, AccessRight.PaymentTrackingJupiter)]
public sealed class PayTrackingController(IPayTrackingService tracking, IPayClinicService clinic) : ApiControllerBase
{
    /// <summary>Clinics (SearchCLN) and group companies (PaymentReceiptModule case 16).</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups()
    {
        var (_, companies, clinics) = await clinic.LookupsAsync();
        return new { companies, clinics = clinics.Select(c => new { code = c.Code, name = c.Name }) };
    }

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<PayTrackingResult> Search(PayTrackingRequest request) => tracking.SearchAsync(request);
}
