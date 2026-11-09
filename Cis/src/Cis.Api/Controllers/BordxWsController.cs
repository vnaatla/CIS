using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Bordx;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim Bordx > Worksheet > Insert Worksheet (frmWSBordx).</summary>
[Route("api/claim-bordx/insert-worksheet")]
[RequireAccess(AccessRight.BordxInsert)]
public sealed class BordxWsController(IBordxWsService bordx) : ApiControllerBase
{
    /// <summary>Group companies (name|code), treatment codes, panel clinics.</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups()
    {
        var (groups, treatments, clinics) = await bordx.LookupsAsync();
        return new
        {
            groups = groups.Select(g => new { code = g.Code, name = g.Name }), treatments = treatments.Select(t => new { code = t.Code, name = t.Name }),
            clinics = clinics.Select(c => new { code = c.Code, name = c.Name }),
        };
    }

    /// <summary>The group company's departments and subsidiaries.</summary>
    [HttpGet("company")]
    public async Task<object> Company([FromQuery] string? groupCompany)
    {
        var (departments, subsidiaries) = await bordx.CompanyAsync(groupCompany);
        return new { departments, subsidiaries };
    }

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<BordxWsResult> Search(BordxWsRequest request) => bordx.SearchAsync(request);

    /// <summary>Bordx: creates the bordereaux for the ticked cases.</summary>
    [HttpPost("bordx")]
    public async Task<object> Bordx(BordxWsSaveRequest request)
    {
        var no = await bordx.BordxAsync(request, UserCode);
        return new { bordxNo = no, message = BordxWsService.Saved + no };
    }
}
