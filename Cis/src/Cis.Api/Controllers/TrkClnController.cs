using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Tracking > Clinic (FrmTracking).</summary>
[Route("api/tracking/clinic")]
[RequireAccess(AccessRight.ClinicTracking)]
public sealed class TrkClnController(ITrkClnService tracking) : ApiControllerBase
{
    /// <summary>Clinic groups, cities and states.</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups()
    {
        var (groups, cities, states) = await tracking.LookupsAsync();
        return new { groups, cities, states };
    }

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(TrkClnRequest request) => tracking.SearchAsync(request);
}
