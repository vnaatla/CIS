using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Bordx;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim Bordx > Worksheet > Insert Invoice (frmWSBordx2).</summary>
[Route("api/claim-bordx/insert-invoice")]
[RequireAccess(AccessRight.BordxInsert)]
public sealed class BordxInvController(IBordxInvService bordx) : ApiControllerBase
{
    /// <summary>Group companies and payors.</summary>
    [HttpGet("lookups")]
    public Task<object> Lookups() => bordx.LookupsAsync();

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<BordxInvResult> Search(BordxInvSearchRequest request) => bordx.SearchAsync(request);

    /// <summary>Bordx: the ticked invoices' claims go into a new bordereaux.</summary>
    [HttpPost("bordx")]
    public async Task<object> Bordx(BordxInvSaveRequest request)
    {
        var no = await bordx.BordxAsync(request, UserCode);
        return new { bordxNo = no, message = BordxInvService.Saved + no };
    }
}
