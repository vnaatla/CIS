using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Mco;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>MCO Fees > MCO Invoice GST (frmMCOFeeGST).</summary>
[Route("api/mco-fees/invoice-gst")]
[RequireAccess(AccessRight.McoInvoice)]
public sealed class McoGstController(IMcoGstService gst, IMcoListService lists) : ApiControllerBase
{
    /// <summary>Payors and group companies.</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups()
    {
        var (payors, groups) = await lists.LookupsAsync();
        return new { payors = payors.Select(p => new { code = p.Code, name = p.Name }), groups = groups.Select(g => new { code = g.Code, name = g.Name }) };
    }

    /// <summary>The group company's subsidiaries.</summary>
    [HttpGet("subsidiaries")]
    public Task<List<string>> Subsidiaries([FromQuery] string? groupCompany) => gst.SubsidiariesAsync(groupCompany);

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<McoGstResult> Search(McoGstRequest request) => gst.SearchAsync(request);

    /// <summary>Print / Print Preview (Excel: invoice + details).</summary>
    [HttpPost("print")]
    public async Task<FileContentResult> Print(McoGstRequest request)
    {
        var file = await gst.PrintAsync(request, UserCode);
        return File(file.Content, file.ContentType, file.FileName);
    }
}
