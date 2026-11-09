using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Bordx;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim Bordx > ClaimBordxDate (FrmBordxDate).</summary>
[Route("api/claim-bordx/bordx-date")]
[RequireAccess(AccessRight.BordxInsert)]
public sealed class BordxDateController(IBordxDateService dates) : ApiControllerBase
{
    /// <summary>Search.</summary>
    [HttpGet]
    public async Task<object> Find([FromQuery] string? bordxNo)
    {
        var (rows, posted) = await dates.FindAsync(bordxNo);
        return new { rows, posted, message = posted ? BordxDateService.Posted : null };
    }

    /// <summary>Update.</summary>
    [HttpPost("update")]
    public async Task<object> Update(BordxDateUpdateRequest request) => new { message = await dates.UpdateAsync(request, UserCode) };
}
