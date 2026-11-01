using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Membership;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Membership > MBM Adjustment (FrmMBMCLNAdj).</summary>
[Route("api/membership/adjustment")]
[RequireAccess(AccessRight.MbmAdjustment)]
public sealed class MbmAdjustmentController(IMbmAdjustmentService adjustment) : ApiControllerBase
{
    /// <summary>Card types and relationships.</summary>
    [HttpGet("setup")]
    public Task<AdjSetup> Setup() => adjustment.SetupAsync();

    /// <summary>Show: the member, its covers and the desktop's status notice.</summary>
    [HttpGet]
    public Task<AdjMember> Find([FromQuery] string? member) => adjustment.FindAsync(member);

    /// <summary>Prev / Next member number.</summary>
    [HttpGet("step")]
    public async Task<object> Step([FromQuery] string? member, [FromQuery] bool next) => new { member = await adjustment.StepAsync(member, next) };

    /// <summary>Adjustment history tab.</summary>
    [HttpGet("amend-history")]
    public Task<List<Dictionary<string, object?>>> AmendHistory([FromQuery] string? member) => adjustment.AmendHistoryAsync(member);

    /// <summary>MBM history tab (by IC).</summary>
    [HttpGet("mbm-history")]
    public Task<List<Dictionary<string, object?>>> MbmHistory([FromQuery] string? ic) => adjustment.MbmHistoryAsync(ic);

    /// <summary>New premium / MCO for a changed policy period.</summary>
    [HttpPost("fees")]
    public async Task<object> Fees(AdjFeesRequest request)
    {
        var (fees, warnings) = await adjustment.FeesAsync(request);
        return new { fees, warnings };
    }

    /// <summary>Save the endorsement.</summary>
    [HttpPost]
    public async Task<object> Save(AdjSaveRequest request) => new { message = await adjustment.SaveAsync(request, UserCode) };
}
