using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.CodeTables;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>
/// The simple maintenance tabs (Check Item, Calculation, Case, Dental / Benefit categories): search, add, edit, delete. The class accepts any
/// maintenance right; the service checks the right of the table asked for.
/// </summary>
[Route("api/maintenance/code-tables")]
[RequireAccess(AccessRight.MaintenanceEntity, AccessRight.MaintenancePlan, AccessRight.MaintenanceCase, AccessRight.MaintenanceCalculation,
    AccessRight.MaintenanceBenefitMember, AccessRight.MaintenanceBenefitClinic, AccessRight.MaintenanceCheckItem)]
public sealed class MaintenanceCodeTablesController(ICodeTableService codeTables) : ApiControllerBase
{
    /// <summary>Fields of the table (labels, lengths, mandatory, searchable).</summary>
    [HttpGet("{key}")]
    public CodeTableView Describe(string key) => codeTables.Describe(key, UserAccess);

    /// <summary>Search: query string = searchable column values; none = all rows.</summary>
    [HttpGet("{key}/rows")]
    public Task<List<Dictionary<string, object?>>> Search(string key) =>
        codeTables.SearchAsync(key, Request.Query.ToDictionary(q => q.Key, q => q.Value.ToString(), StringComparer.OrdinalIgnoreCase), UserAccess);

    /// <summary>Options of a choice field.</summary>
    [HttpGet("{key}/choices/{column}")]
    public Task<List<LookupOption>> Choices(string key, string column) => codeTables.ChoicesAsync(key, column, UserAccess);

    /// <summary>Add + Save.</summary>
    [HttpPost("{key}/rows")]
    public Task<CodeTableResult> Create(string key, CodeTableSaveRequest request) => codeTables.CreateAsync(key, request, UserCode, UserAccess);

    /// <summary>Edit + Save.</summary>
    [HttpPut("{key}/rows/{code}")]
    public Task<CodeTableResult> Update(string key, string code, CodeTableSaveRequest request) => codeTables.UpdateAsync(key, code, request, UserCode, UserAccess);

    /// <summary>Delete.</summary>
    [HttpDelete("{key}/rows/{code}")]
    public Task<CodeTableResult> Delete(string key, string code) => codeTables.DeleteAsync(key, code, UserAccess);
}
