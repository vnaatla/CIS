using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Receipt;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Receipt > Fr FFS (FrmReceiptFFS).</summary>
[Route("api/receipt/ffs")]
[RequireAccess(AccessRight.ReceiptFfs)]
public sealed class RcpFfsController(IRcpFfsService receipt, IRcpPayService payor) : ApiControllerBase
{
    /// <summary>Group companies.</summary>
    [HttpGet("lookups")]
    public Task<object> Lookups() => payor.LookupsAsync();

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<RcpFfsResult> Search(RcpFfsSearchRequest request) => receipt.SearchAsync(request);

    /// <summary>The rows of the receipt just saved.</summary>
    [HttpGet("receipt/{receiptNo}")]
    public Task<RcpFfsResult> Receipt(string receiptNo) => receipt.SearchAsync(new RcpFfsSearchRequest { Scan = false }, receiptNo);

    /// <summary>Save the receipt for the ticked bordereaux.</summary>
    [HttpPost("save")]
    public async Task<object> Save(RcpFfsSaveRequest request) => new { message = await receipt.SaveAsync(request, UserCode) };

    /// <summary>Update the selected receipt.</summary>
    [HttpPost("update")]
    public async Task<object> Update(RcpFfsSaveRequest request) => new { message = await receipt.UpdateAsync(request, UserCode) };
}
