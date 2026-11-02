using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Other > Invoice Update (frmUpdateCLNInvoice).</summary>
[Route("api/invoice/update")]
[RequireAccess(AccessRight.InvoiceAdjustment)]
public sealed class InvoiceUpdateController(IInvoiceUpdateService invoices) : ApiControllerBase
{
    /// <summary>Panel clinics and group companies.</summary>
    [HttpGet("setup")]
    public Task<InvUpdateSetup> Setup() => invoices.SetupAsync();

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(InvUpdateFilters filters) => invoices.SearchAsync(filters);

    /// <summary>Print: the list as Excel.</summary>
    [HttpPost("export")]
    public async Task<FileContentResult> Export(InvUpdateFilters filters)
    {
        var file = await invoices.ExportAsync(filters);
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Update the checked cases' invoice number and date.</summary>
    [HttpPost]
    public async Task<object> Save(InvUpdateRequest request)
    {
        var (message, rows) = await invoices.SaveAsync(request, UserCode);
        return new { message, rows };
    }
}
