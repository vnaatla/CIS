using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Other > MMA/CPT (frmMMASchedule, frmShowInvestigation).</summary>
[Route("api/invoice/mma-cpt")]
[RequireAccess(AccessRight.MmaCpt)]
public sealed class MmaScheduleController(IMmaScheduleService mma) : ApiControllerBase
{
    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<MmaRow>> Search(MmaFilters filters) => mma.SearchAsync(filters);

    /// <summary>Print: the list as Excel.</summary>
    [HttpPost("export")]
    public async Task<FileContentResult> Export(MmaFilters filters)
    {
        var file = await mma.ExportAsync(filters);
        return File(file.Content, file.ContentType, file.FileName);
    }
}
