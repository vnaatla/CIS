using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Bordx;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim Bordx > Worksheet > Remove Worksheet (FrmBordxRevise).</summary>
[Route("api/claim-bordx/remove-worksheet")]
[RequireAccess(AccessRight.BordxRemove)]
public sealed class BordxRemoveController(IBordxRemoveService remove) : ApiControllerBase
{
    /// <summary>Go: the bordereaux's worksheets.</summary>
    [HttpGet]
    public Task<BordxRemoveResult> Find([FromQuery] string? bordxNo) => remove.FindAsync(bordxNo);

    /// <summary>Reject: removes the ticked worksheets.</summary>
    [HttpPost("remove")]
    public async Task<object> Remove(BordxRemoveRequest request) => new { message = await remove.RemoveAsync(request, UserCode) };
}
