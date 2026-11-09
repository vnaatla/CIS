using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Receipt;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Receipt > Fr MNRB (FrmReceiptMNRB).</summary>
[Route("api/receipt/mnrb")]
[RequireAccess(AccessRight.ReceiptPayor)]
public sealed class RcpMnrbController(IRcpMnrbService receipt) : ApiControllerBase
{
    /// <summary>The MNRB payors.</summary>
    [HttpGet("payors")]
    public async Task<object> Payors() => (await receipt.PayorsAsync()).Select(p => new { code = p.Code, name = p.Name });

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(RcpMnrbSearchRequest request) => receipt.SearchAsync(request);

    /// <summary>The receipt of a row (edit mode).</summary>
    [HttpGet("receipt")]
    public Task<Dictionary<string, object?>> Receipt([FromQuery] string bordxNo, [FromQuery] string? casNumber) => receipt.ReceiptAsync(bordxNo, casNumber);

    /// <summary>Save the receipt for the ticked rows.</summary>
    [HttpPost("save")]
    public async Task<object> Save(RcpMnrbSaveRequest request) => new { message = await receipt.SaveAsync(request, UserCode) };

    /// <summary>Update the selected row's receipt.</summary>
    [HttpPost("update")]
    public async Task<object> Update(RcpMnrbSaveRequest request) => new { message = await receipt.UpdateAsync(request, UserCode) };
}
