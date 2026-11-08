using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Mco;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>MCO Fees > MCO Invoice (frmMCOInvoiceNew).</summary>
[Route("api/mco-fees/invoice")]
[RequireAccess(AccessRight.McoInvoice)]
public sealed class McoInvoiceController(IMcoInvoiceService invoices, IMcoListService lists) : ApiControllerBase
{
    /// <summary>Payors and group companies (as the fee listing).</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups()
    {
        var (payors, groups) = await lists.LookupsAsync();
        return new { payors = payors.Select(p => new { code = p.Code, name = p.Name }), groups = groups.Select(g => new { code = g.Code, name = g.Name }) };
    }

    /// <summary>The group company's departments.</summary>
    [HttpGet("departments")]
    public Task<List<string>> Departments([FromQuery] string? groupCompany) => invoices.DepartmentsAsync(groupCompany);

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<McoInvoiceResult> Search(McoInvoiceRequest request) => invoices.SearchAsync(request);

    /// <summary>Print / Print Preview (Excel). Print numbers and stamps a new invoice.</summary>
    [HttpPost("print")]
    public async Task<FileContentResult> Print(McoInvoiceRequest request)
    {
        var file = await invoices.PrintAsync(request, UserCode);
        return File(file.Content, file.ContentType, file.FileName);
    }
}
