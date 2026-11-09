using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Bordx;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim Bordx > Bordx Delivery: Send to Grp. Company (payor), Returned Fr Grp. Company (returned), Send to A/C (ac), Send to MNRB (mnrb).</summary>
[Route("api/claim-bordx/delivery")]
public sealed class BordxDeliveryController(IBordxDeliveryService delivery) : ApiControllerBase
{
    /// <summary>Group companies and payors.</summary>
    [HttpGet("lookups")]
    [RequireAccess(AccessRight.BordxSendToPayor, AccessRight.BordxReturnedFromPayor, AccessRight.BordxSendToAc)]
    public async Task<object> Lookups()
    {
        var (companies, payors) = await delivery.LookupsAsync();
        return new { companies, payors = payors.Select(p => new { code = p.Code, name = p.Name }) };
    }

    /// <summary>Send to Grp. Company / Send to MNRB.</summary>
    [HttpPost("payor/search")]
    [RequireAccess(AccessRight.BordxSendToPayor)]
    public Task<BordxDeliveryResult> SearchPayor(BordxDeliveryRequest request) => delivery.SearchAsync("payor", request);

    /// <summary>Update.</summary>
    [HttpPost("payor/update")]
    [RequireAccess(AccessRight.BordxSendToPayor)]
    public async Task<object> UpdatePayor(BordxDeliveryUpdateRequest request) => new { message = await delivery.UpdateAsync("payor", request) };

    /// <summary>Send to MNRB.</summary>
    [HttpPost("mnrb/search")]
    [RequireAccess(AccessRight.BordxSendToPayor)]
    public Task<BordxDeliveryResult> SearchMnrb(BordxDeliveryRequest request) => delivery.SearchAsync("mnrb", request);

    /// <summary>Update.</summary>
    [HttpPost("mnrb/update")]
    [RequireAccess(AccessRight.BordxSendToPayor)]
    public async Task<object> UpdateMnrb(BordxDeliveryUpdateRequest request) => new { message = await delivery.UpdateAsync("mnrb", request) };

    /// <summary>Returned Fr Grp. Company.</summary>
    [HttpPost("returned/search")]
    [RequireAccess(AccessRight.BordxReturnedFromPayor)]
    public Task<BordxDeliveryResult> SearchReturned(BordxDeliveryRequest request) => delivery.SearchAsync("returned", request);

    /// <summary>Update.</summary>
    [HttpPost("returned/update")]
    [RequireAccess(AccessRight.BordxReturnedFromPayor)]
    public async Task<object> UpdateReturned(BordxDeliveryUpdateRequest request) => new { message = await delivery.UpdateAsync("returned", request) };

    /// <summary>Send to A/C.</summary>
    [HttpPost("ac/search")]
    [RequireAccess(AccessRight.BordxSendToAc)]
    public Task<BordxDeliveryResult> SearchAc(BordxDeliveryRequest request) => delivery.SearchAsync("ac", request);

    /// <summary>Update.</summary>
    [HttpPost("ac/update")]
    [RequireAccess(AccessRight.BordxSendToAc)]
    public async Task<object> UpdateAc(BordxDeliveryUpdateRequest request) => new { message = await delivery.UpdateAsync("ac", request) };
}
