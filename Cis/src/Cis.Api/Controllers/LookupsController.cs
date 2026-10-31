using Cis.Api.Common;
using Cis.Api.Lookups;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Shared drop-down lists (any signed-in user; each screen's own endpoints carry its access check).</summary>
[Route("api/lookups")]
public sealed class LookupsController(ILookupService lookups) : ApiControllerBase
{
    /// <summary>Payors (CISDB.PAY), optionally of one health type.</summary>
    [HttpGet("payors")]
    public Task<List<LookupOption>> Payors([FromQuery] string? healthType) => lookups.PayorsAsync(healthType);

    /// <summary>Health types (CISDB.HLT).</summary>
    [HttpGet("health-types")]
    public Task<List<LookupOption>> HealthTypes() => lookups.HealthTypesAsync();

    /// <summary>Insured types (CISDB.INS).</summary>
    [HttpGet("insured-types")]
    public Task<List<LookupOption>> InsuredTypes() => lookups.InsuredTypesAsync();

    /// <summary>Group companies (CISDB.GRPCompany), optionally of one payor.</summary>
    [HttpGet("group-companies")]
    public Task<List<LookupOption>> GroupCompanies([FromQuery] string? payor) => lookups.GroupCompaniesAsync(payor);

    /// <summary>Clinics (HISDB.CLN), optionally panel clinics only and / or matching a code or name.</summary>
    [HttpGet("clinics")]
    public Task<List<LookupOption>> Clinics([FromQuery] bool panelOnly, [FromQuery] string? search) => lookups.ClinicsAsync(panelOnly, search);
}
