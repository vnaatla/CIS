using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.Plan;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Maintenance > Plan (frmMaintenancePlan): Plan Code, Premium, Annual Limit, Visit Limit and MCO Fee tabs (HealthType / Insured Type are code-tables).</summary>
[Route("api/maintenance/plan")]
[RequireAccess(AccessRight.MaintenancePlan)]
public sealed class MaintenancePlanController(IPlanCodeService plans, IPlanPremiumService premiums, IPlanAnnualLimitService annual, IPlanVisitLimitService visits, IPlanMcoService mcos) : ApiControllerBase
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

    /// <summary>Premium combos.</summary>
    [HttpGet("premiums/setup")]
    public Task<PremiumSetup> PremiumSetup() => premiums.SetupAsync();

    /// <summary>Premium search.</summary>
    [HttpPost("premiums/search")]
    public Task<List<Dictionary<string, object?>>> Premiums(PremiumRequest request) => premiums.SearchAsync(request.Premium);

    /// <summary>Band amounts of insured type + plan (row click).</summary>
    [HttpGet("premiums/bands")]
    public Task<List<PremiumBand>> PremiumBands([FromQuery] string? insCode, [FromQuery] string? plnCode) => premiums.BandsAsync(insCode, plnCode);

    /// <summary>Premium Add + Save (seven age-band rows).</summary>
    [HttpPost("premiums")]
    public async Task<object> CreatePremium(PremiumRequest request) => new { message = await premiums.SaveAsync(request.Premium, UserCode) };

    /// <summary>Premium Delete (one PRM row; the keys as shown for the membership check).</summary>
    [HttpPost("premiums/{prmCode:int}/delete")]
    public async Task<object> DeletePremium(int prmCode, PremiumRequest request) => new { message = await premiums.DeleteAsync(prmCode, request.Premium) };

    /// <summary>Annual Limit combos.</summary>
    [HttpGet("annual-limits/setup")]
    public Task<AnnualLimitSetup> AnnualSetup() => annual.SetupAsync();

    /// <summary>Annual Limit search.</summary>
    [HttpPost("annual-limits/search")]
    public Task<List<Dictionary<string, object?>>> AnnualLimits(AnnualLimitRequest request) => annual.SearchAsync(request.Limit);

    /// <summary>Annual Limit Add + Save.</summary>
    [HttpPost("annual-limits")]
    public async Task<object> CreateAnnualLimit(AnnualLimitRequest request) => new { message = await annual.SaveAsync(request.Limit, false, UserCode) };

    /// <summary>Annual Limit Edit + Save (Limit.Index).</summary>
    [HttpPut("annual-limits")]
    public async Task<object> UpdateAnnualLimit(AnnualLimitRequest request) => new { message = await annual.SaveAsync(request.Limit, true, UserCode) };

    /// <summary>Annual Limit Delete (Limit.Index; plan as shown for the member check).</summary>
    [HttpPost("annual-limits/delete")]
    public async Task<object> DeleteAnnualLimit(AnnualLimitRequest request) => new { message = await annual.DeleteAsync(request.Limit) };

    /// <summary>Visit Limit search (combos: annual-limits/setup).</summary>
    [HttpPost("visit-limits/search")]
    public Task<List<Dictionary<string, object?>>> VisitLimits(VisitLimitRequest request) => visits.SearchAsync(request.Limit);

    /// <summary>Visit Limit Add + Save.</summary>
    [HttpPost("visit-limits")]
    public async Task<object> CreateVisitLimit(VisitLimitRequest request) => new { message = await visits.SaveAsync(request.Limit, false, UserCode) };

    /// <summary>Visit Limit Edit + Save (Limit.Index).</summary>
    [HttpPut("visit-limits")]
    public async Task<object> UpdateVisitLimit(VisitLimitRequest request) => new { message = await visits.SaveAsync(request.Limit, true, UserCode) };

    /// <summary>Visit Limit Delete (Limit.Index; plan as shown for the member check).</summary>
    [HttpPost("visit-limits/delete")]
    public async Task<object> DeleteVisitLimit(VisitLimitRequest request) => new { message = await visits.DeleteAsync(request.Limit) };

    /// <summary>MCO Fee combos.</summary>
    [HttpGet("mco/setup")]
    public Task<McoSetup> McoSetup() => mcos.SetupAsync();

    /// <summary>MCO Fee search.</summary>
    [HttpPost("mco/search")]
    public Task<List<Dictionary<string, object?>>> Mco(McoRequest request) => mcos.SearchAsync(request.Mco);

    /// <summary>MCO Fee Add + Save (messages in order: per-type refusals, then the saved text).</summary>
    [HttpPost("mco")]
    public async Task<object> CreateMco(McoRequest request) => new { messages = await mcos.SaveAsync(request.Mco, UserCode) };

    /// <summary>MCO Fee Delete of the selected row.</summary>
    [HttpDelete("mco/{index:int}")]
    public async Task<object> DeleteMco(int index, [FromQuery] string? healthType) => new { message = await mcos.DeleteAsync(index, healthType) };
}
