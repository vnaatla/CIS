using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Receipt;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Receipt > Fr Member (FrmReceiptMember).</summary>
[Route("api/receipt/member")]
[RequireAccess(AccessRight.ReceiptMember)]
public sealed class RcpMbmController(IRcpMbmService receipt) : ApiControllerBase
{
    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(RcpMbmSearchRequest request) => receipt.SearchAsync(request);

    /// <summary>The receipt of a row (edit mode).</summary>
    [HttpGet("receipt")]
    public Task<Dictionary<string, object?>> Receipt([FromQuery] string? index) => receipt.ReceiptAsync(index);

    /// <summary>Save the receipt for the ticked cases.</summary>
    [HttpPost("save")]
    public async Task<object> Save(RcpMbmSaveRequest request) => new { message = await receipt.SaveAsync(request, UserCode) };

    /// <summary>Update the selected receipt.</summary>
    [HttpPost("update")]
    public async Task<object> Update(RcpMbmSaveRequest request) => new { message = await receipt.UpdateAsync(request, UserCode) };
}
