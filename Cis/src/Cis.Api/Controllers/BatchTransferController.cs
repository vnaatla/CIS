using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Clinic > Batch Transfer (FrmInvoiceBatchNoUpd).</summary>
[Route("api/invoice/batch-transfer")]
[RequireAccess(AccessRight.InvoiceAdjustment)]
public sealed class BatchTransferController(IBatchTransferService transfer) : ApiControllerBase
{
    /// <summary>A batch's cases, invoice, amount and clinic.</summary>
    [HttpGet("{batchNo}")]
    public Task<BatchInfo> Batch(string batchNo) => transfer.BatchAsync(batchNo);

    /// <summary>Transfer.</summary>
    [HttpPost]
    public async Task<object> Save(BatchTransferRequest request) => new { message = await transfer.SaveAsync(request, UserCode) };
}
