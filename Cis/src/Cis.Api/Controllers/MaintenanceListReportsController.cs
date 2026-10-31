using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Documents;
using Cis.Application.Maintenance.ListReports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>The maintenance screens' Print buttons (the desktop's DataReports). The service checks the right of the report asked for.</summary>
[Route("api/maintenance/list-reports")]
[RequireAccess(AccessRight.MaintenanceEntity, AccessRight.MaintenancePlan, AccessRight.MaintenanceCase, AccessRight.MaintenanceCalculation,
    AccessRight.MaintenanceBenefitMember, AccessRight.MaintenanceBenefitClinic, AccessRight.MaintenanceCheckItem)]
public sealed class MaintenanceListReportsController(IListReportService reports) : ApiControllerBase
{
    /// <summary>The report as a file (PDF by default; <c>format=excel|csv</c> for the other types).</summary>
    [HttpGet("{key}")]
    public async Task<FileContentResult> Print(string key, [FromQuery] string? code, [FromQuery] ReportFileType format = ReportFileType.Pdf)
    {
        var file = await reports.PrintAsync(key, code, format, UserAccess);
        return File(file.Content, file.ContentType, file.FileName);
    }
}
