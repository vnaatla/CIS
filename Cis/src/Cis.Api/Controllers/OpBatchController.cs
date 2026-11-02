using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Other > OP Batch/Invoice (frmUpdateInvoiceBatch). The combos come from <c>api/invoice/update/setup</c>.</summary>
[Route("api/invoice/op-batch")]
[RequireAccess(AccessRight.InvoiceAdjustment)]
public sealed class OpBatchController(IOpBatchService batches) : ApiControllerBase
{
    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<OpBatchList> Search(OpBatchFilters filters) => batches.SearchAsync(filters);

    /// <summary>Print: the list as Excel.</summary>
    [HttpPost("export")]
    public async Task<FileContentResult> Export(OpBatchFilters filters)
    {
        var file = await batches.ExportAsync(filters);
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Update: OP batch and / or invoice number of the checked cases.</summary>
    [HttpPost]
    public async Task<object> Save(OpBatchRequest request)
    {
        var (message, batchNo, invoiceNo) = await batches.SaveAsync(request, UserCode);
        return new { message, batchNo, invoiceNo };
    }
}
