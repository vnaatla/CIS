using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.Plan;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Maintenance > Plan (frmMaintenancePlan): Plan Code tab (HealthType / Insured Type are code-tables).</summary>
[Route("api/maintenance/plan")]
[RequireAccess(AccessRight.MaintenancePlan)]
public sealed class MaintenancePlanController(IPlanCodeService plans) : ApiControllerBase
{
    /// <summary>Plan Code combos.</summary>
    [HttpGet("codes/setup")]
    public Task<PlanCodeSetup> Setup() => plans.SetupAsync();

    /// <summary>CIS payors of a health type.</summary>
    [HttpGet("codes/payors")]
    public Task<List<LookupOption>> Payors([FromQuery] string? healthType) => plans.PayorsAsync(healthType);

    /// <summary>Search (the boxes are the filters).</summary>
    [HttpPost("codes/search")]
    public Task<List<Dictionary<string, object?>>> Search(PlanCodeRequest request) => plans.SearchAsync(request.Plan);

    /// <summary>One plan.</summary>
    [HttpGet("codes/{code}")]
    public Task<Dictionary<string, object?>> Get(string code) => plans.GetAsync(code);

    /// <summary>Add + Save (health type N: one to six generated plans).</summary>
    [HttpPost("codes")]
    public async Task<object> Create(PlanCodeRequest request)
    {
        var (message, codes) = await plans.SaveAsync(request.Plan, UserCode);
        return new { message, codes };
    }

    /// <summary>Delete.</summary>
    [HttpDelete("codes/{code}")]
    public async Task<object> Delete(string code) => new { message = await plans.DeleteAsync(code) };
}
