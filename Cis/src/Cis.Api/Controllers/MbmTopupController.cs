using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Membership;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Membership > MBM Topup Adj (FrmMBMTopupLmt).</summary>
[Route("api/membership/topup")]
[RequireAccess(AccessRight.MbmTopup)]
public sealed class MbmTopupController(IMbmTopupService topup) : ApiControllerBase
{
    /// <summary>Go: the member ("NUMBER-COVID") with the GP or SP balance.</summary>
    [HttpGet]
    public Task<TopupMember> Find([FromQuery] string? member, [FromQuery] bool gp = true) => topup.FindAsync(member, gp);

    /// <summary>Next / Previous member number.</summary>
    [HttpGet("step")]
    public async Task<object> Step([FromQuery] string? member, [FromQuery] bool next) => new { member = await topup.StepAsync(member, next) };

    /// <summary>Save.</summary>
    [HttpPost]
    public async Task<object> Save(TopupSaveRequest request) => new { message = await topup.SaveAsync(request, UserCode) };
}
