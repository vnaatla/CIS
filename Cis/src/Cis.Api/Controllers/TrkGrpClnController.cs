using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Tracking > Grp. Com. Clinic (FrmTracking2).</summary>
[Route("api/tracking/grp-company-clinic")]
[RequireAccess(AccessRight.ClinicTracking)]
public sealed class TrkGrpClnController(ITrkGrpClnService tracking) : ApiControllerBase
{
    /// <summary>Group company and clinic names.</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups()
    {
        var (companies, clinics) = await tracking.LookupsAsync();
        return new { companies, clinics };
    }

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(TrkGrpClnRequest request) => tracking.SearchAsync(request);
}
