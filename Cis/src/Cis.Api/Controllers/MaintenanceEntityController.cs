using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Common;
using Cis.Application.Maintenance.Entity;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Maintenance > Entity (frmMaintenanceEntity): all five tabs and frmSearchCLN.</summary>
[Route("api/maintenance/entity")]
[RequireAccess(AccessRight.MaintenanceEntity)]
public sealed class MaintenanceEntityController(IEntityGroupDoctorService entity, IEntityClinicService clinics, IEntityGrpComAccessService links) : ApiControllerBase
{
    /// <summary>Clinic groups (code / name contain).</summary>
    [HttpGet("groups")]
    public async Task<PagedResult<EntityListRow>> Groups([FromQuery] string? code, [FromQuery] string? name, [FromQuery] PageQuery paging) =>
        Paging.Apply(await entity.SearchGroupsAsync(code, name), paging);

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
    public async Task<PagedResult<EntityListRow>> Doctors([FromQuery] string? code, [FromQuery] string? name, [FromQuery] PageQuery paging) =>
        Paging.Apply(await entity.SearchDoctorsAsync(code, name), paging);

    /// <summary>Add + Save (code generated).</summary>
    [HttpPost("doctors")]
    public async Task<object> CreateDoctor(DoctorSaveRequest request) => new { message = await entity.SaveDoctorAsync(request, false, UserCode) };

    /// <summary>Edit + Save.</summary>
    [HttpPut("doctors")]
    public async Task<object> UpdateDoctor(DoctorSaveRequest request) => new { message = await entity.SaveDoctorAsync(request, true, UserCode) };

    /// <summary>Delete.</summary>
    [HttpDelete("doctors/{code}")]
    public async Task<object> DeleteDoctor(string code) => new { message = await entity.DeleteDoctorAsync(code) };

    /// <summary>Clinic Info combos.</summary>
    [HttpGet("clinics/setup")]
    public Task<ClinicSetup> ClinicSetup() => clinics.SetupAsync();

    /// <summary>Clinic Info search (the boxes are the filters).</summary>
    [HttpPost("clinics/search")]
    public async Task<PagedResult<Dictionary<string, object?>>> Clinics(ClinicRequest request, [FromQuery] PageQuery paging) =>
        Paging.Apply(await clinics.SearchAsync(request.Clinic), paging);

    /// <summary>Clinic Info print: the list as Excel.</summary>
    [HttpPost("clinics/export")]
    public async Task<FileContentResult> ExportClinics(ClinicRequest request)
    {
        var file = await clinics.ExportAsync(request.Clinic);
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Add + Save (code generated).</summary>
    [HttpPost("clinics")]
    public async Task<object> CreateClinic(ClinicRequest request)
    {
        var (message, code) = await clinics.SaveAsync(request.Clinic, false, UserCode);
        return new { message, code };
    }

    /// <summary>Edit + Save.</summary>
    [HttpPut("clinics")]
    public async Task<object> UpdateClinic(ClinicRequest request)
    {
        var (message, code) = await clinics.SaveAsync(request.Clinic, true, UserCode);
        return new { message, code };
    }

    /// <summary>Delete (name as on screen, for the desktop's first check).</summary>
    [HttpDelete("clinics/{code}")]
    public async Task<object> DeleteClinic(string code, [FromQuery] string? name) => new { message = await clinics.DeleteAsync(code, name) };

    /// <summary>Group Company Clinic combos (Check1).</summary>
    [HttpGet("group-company-clinics/lookups")]
    public Task<GroupCompanyClinicLookups> GroupCompanyClinicLookups() => links.LookupsAsync();

    /// <summary>Group Company Clinic search (= group company / = clinic; capped list + total).</summary>
    [HttpGet("group-company-clinics")]
    public Task<GroupCompanyClinicPage> GroupCompanyClinics([FromQuery] string? groupCompany, [FromQuery] string? clinic, [FromQuery] int page = 0,
        [FromQuery] int pageSize = 50, [FromQuery] string? sortKey = null, [FromQuery] bool sortDesc = false) =>
        links.SearchAsync(groupCompany, clinic, page, pageSize, sortKey, sortDesc);

    /// <summary>Add + Save.</summary>
    [HttpPost("group-company-clinics")]
    public async Task<object> CreateGroupCompanyClinic(GroupCompanyClinicSaveRequest request) => new { message = await links.SaveAsync(request, false, UserCode) };

    /// <summary>Edit + Save (the selected pair in Old*).</summary>
    [HttpPut("group-company-clinics")]
    public async Task<object> UpdateGroupCompanyClinic(GroupCompanyClinicSaveRequest request) => new { message = await links.SaveAsync(request, true, UserCode) };

    /// <summary>Delete the selected pair.</summary>
    [HttpDelete("group-company-clinics")]
    public async Task<object> DeleteGroupCompanyClinic([FromQuery] string? groupCompany, [FromQuery] string? clinic) =>
        new { message = await links.DeleteAsync(groupCompany, clinic) };

    /// <summary>Clinic Access: the latest 100 rows.</summary>
    [HttpGet("clinic-access")]
    public Task<List<ClinicAccess>> ClinicAccessList() => links.AccessListAsync();

    /// <summary>Clinic Access: Add + Save.</summary>
    [HttpPost("clinic-access")]
    public async Task<object> CreateClinicAccess(ClinicAccessSaveRequest request) => new { message = await links.SaveAccessAsync(request, UserCode) };

    /// <summary>frmSearchCLN: clinics whose name contains the text.</summary>
    [HttpGet("clinic-search")]
    public Task<List<ClinicFound>> FindClinics([FromQuery] string? name) => links.FindClinicsAsync(name);
}
