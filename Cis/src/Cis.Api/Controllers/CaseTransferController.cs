using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Membership;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Other > Case Transfer (frmMembershipTransfer).</summary>
[Route("api/invoice/case-transfer")]
[RequireAccess(AccessRight.CaseTransfer)]
public sealed class CaseTransferController(ICaseTransferService transfer) : ApiControllerBase
{
    /// <summary>Go: the case and its current member.</summary>
    [HttpGet("{casNumber}")]
    public Task<TransferCase> Find(string casNumber) => transfer.FindAsync(casNumber);

    /// <summary>Search: the transfer member's name.</summary>
    [HttpGet("member")]
    public async Task<object> Member([FromQuery] string? mbmNumber, [FromQuery] string? covId) => new { name = await transfer.MemberAsync(mbmNumber, covId) };

    /// <summary>Transfer.</summary>
    [HttpPost]
    public async Task<object> Save(CaseTransferRequest request) => new { message = await transfer.SaveAsync(request) };
}
