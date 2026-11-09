using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Receipt;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Receipt > Fr Grp Company (FrmReceiptPayor).</summary>
[Route("api/receipt/payor")]
[RequireAccess(AccessRight.ReceiptPayor)]
public sealed class RcpPayController(IRcpPayService receipt) : ApiControllerBase
{
    /// <summary>Group companies and CIS payors.</summary>
    [HttpGet("lookups")]
    public Task<object> Lookups() => receipt.LookupsAsync();

    /// <summary>The payor's insured clients.</summary>
    [HttpGet("clients")]
    public Task<List<string>> Clients([FromQuery] string payor) => receipt.ClientsAsync(payor);

    /// <summary>The group company's subsidiaries.</summary>
    [HttpGet("subsidiaries")]
    public Task<List<string>> Subsidiaries([FromQuery] string company) => receipt.SubsidiariesAsync(company);

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(RcpPaySearchRequest request) => receipt.SearchAsync(request);

    /// <summary>The rows of the receipt just saved.</summary>
    [HttpPost("relist")]
    public Task<List<Dictionary<string, object?>>> Relist(RcpPaySearchRequest request, [FromQuery] string receiptNo) =>
        receipt.SearchAsync(request with { Scan = false }, receiptNo);

    /// <summary>The receipt of a row (edit mode).</summary>
    [HttpGet("receipt")]
    public Task<Dictionary<string, object?>> Receipt([FromQuery] string bordxNo, [FromQuery] string? casNumber) => receipt.ReceiptAsync(bordxNo, casNumber);

    /// <summary>Save the receipt for the ticked rows.</summary>
    [HttpPost("save")]
    public async Task<object> Save(RcpPaySaveRequest request) => new { message = await receipt.SaveAsync(request, UserCode) };

    /// <summary>Update the selected row's receipt.</summary>
    [HttpPost("update")]
    public async Task<object> Update(RcpPaySaveRequest request) => new { message = await receipt.UpdateAsync(request, UserCode) };
}
