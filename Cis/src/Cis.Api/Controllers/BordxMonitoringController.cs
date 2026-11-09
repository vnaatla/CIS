using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Bordx;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim Bordx > Monitoring: Bordx Control (FrmBordxControl), Outstanding Bordx (FrmOSBordx), Bordx &amp; Worksheet (FrmBordxWks).</summary>
[Route("api/claim-bordx/monitoring")]
public sealed class BordxMonitoringController(IBordxMonitoringService monitoring) : ApiControllerBase
{
    /// <summary>Group companies, payors and (for a company) its subsidiaries.</summary>
    [HttpGet("lookups")]
    [RequireAccess(AccessRight.BordxControl, AccessRight.BordxOutstanding, AccessRight.BordxWorksheet)]
    public async Task<object> Lookups([FromQuery] string? company)
    {
        var (companies, payors, subsidiaries) = await monitoring.LookupsAsync(company);
        return new { companies, payors = payors.Select(p => new { code = p.Code, name = p.Name }), subsidiaries };
    }

    /// <summary>Bordx Control.</summary>
    [HttpPost("control")]
    [RequireAccess(AccessRight.BordxControl)]
    public Task<BordxListResult> Control(BordxControlRequest request) => monitoring.ControlAsync(request);

    /// <summary>Outstanding Bordx.</summary>
    [HttpGet("outstanding")]
    [RequireAccess(AccessRight.BordxOutstanding)]
    public Task<BordxListResult> Outstanding([FromQuery] string? company, [FromQuery] string? osDate) => monitoring.OutstandingAsync(company, osDate);

    /// <summary>Bordx &amp; Worksheet.</summary>
    [HttpPost("bordx-worksheet")]
    [RequireAccess(AccessRight.BordxWorksheet)]
    public Task<BordxListResult> Worksheets(BordxWksRequest request) => monitoring.WorksheetsAsync(request);
}
