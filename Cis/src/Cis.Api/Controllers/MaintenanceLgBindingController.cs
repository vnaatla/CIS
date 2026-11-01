using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Maintenance.LgTemplates;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Maintenance > LG Maintenance > GL Template Binding (frmLGMaintenance, Tbl_OPGLMaster).</summary>
[Route("api/maintenance/lg-binding")]
[RequireAccess(AccessRight.MaintenanceLg)]
public sealed class MaintenanceLgBindingController(ILgBindingService bindings) : ApiControllerBase
{
    /// <summary>Drop-downs of the tab.</summary>
    [HttpGet("lookups")]
    public Task<LgBindingLookups> Lookups() => bindings.LookupsAsync();

    /// <summary>Plans and pay codes of a group company code.</summary>
    [HttpGet("groups/{groupCode}/plans")]
    public Task<LgGroupPlans> Plans(string groupCode) => bindings.PlansOfGroupAsync(groupCode);

    /// <summary>Search by group company name (required), pay code, plan code.</summary>
    [HttpGet]
    public Task<List<LgBinding>> Search([FromQuery] string? groupName, [FromQuery] string? payCode, [FromQuery] string? planCode) =>
        bindings.SearchAsync(groupName, payCode, planCode);

    /// <summary>Add + Save.</summary>
    [HttpPost]
    public async Task<object> Create(LgBindingSaveRequest request) => new { message = await bindings.CreateAsync(request) };

    /// <summary>Edit + Save (every row of the plan code).</summary>
    [HttpPut]
    public async Task<object> Update(LgBindingSaveRequest request) => new { message = await bindings.UpdateAsync(request) };
}
