using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Maintenance.Entity;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Maintenance > Entity (frmMaintenanceEntity): Clinic Group and Doctor tabs.</summary>
[Route("api/maintenance/entity")]
[RequireAccess(AccessRight.MaintenanceEntity)]
public sealed class MaintenanceEntityController(IEntityGroupDoctorService entity) : ApiControllerBase
{
    /// <summary>Clinic groups (code / name contain).</summary>
    [HttpGet("groups")]
    public Task<List<EntityListRow>> Groups([FromQuery] string? code, [FromQuery] string? name) => entity.SearchGroupsAsync(code, name);

    /// <summary>One clinic group.</summary>
    [HttpGet("groups/{code}")]
    public Task<ClinicGroupDto> Group(string code) => entity.GetGroupAsync(code);

    /// <summary>Add + Save (code generated).</summary>
    [HttpPost("groups")]
    public async Task<object> CreateGroup(ClinicGroupSaveRequest request) => new { message = await entity.SaveGroupAsync(request, false, UserCode) };

    /// <summary>Edit + Save.</summary>
    [HttpPut("groups")]
    public async Task<object> UpdateGroup(ClinicGroupSaveRequest request) => new { message = await entity.SaveGroupAsync(request, true, UserCode) };

    /// <summary>Delete.</summary>
    [HttpDelete("groups/{code}")]
    public async Task<object> DeleteGroup(string code) => new { message = await entity.DeleteGroupAsync(code) };

    /// <summary>Doctors (code / name contain).</summary>
    [HttpGet("doctors")]
    public Task<List<EntityListRow>> Doctors([FromQuery] string? code, [FromQuery] string? name) => entity.SearchDoctorsAsync(code, name);

    /// <summary>Add + Save (code generated).</summary>
    [HttpPost("doctors")]
    public async Task<object> CreateDoctor(DoctorSaveRequest request) => new { message = await entity.SaveDoctorAsync(request, false, UserCode) };

    /// <summary>Edit + Save.</summary>
    [HttpPut("doctors")]
    public async Task<object> UpdateDoctor(DoctorSaveRequest request) => new { message = await entity.SaveDoctorAsync(request, true, UserCode) };

    /// <summary>Delete.</summary>
    [HttpDelete("doctors/{code}")]
    public async Task<object> DeleteDoctor(string code) => new { message = await entity.DeleteDoctorAsync(code) };
}
