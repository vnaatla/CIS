using Cis.Application.Audit;
using Cis.Api.Common;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Legacy action log (HISDB.dbo.tblActionLog, ProjType 'CIS').</summary>
[Route("api/audit")]
public sealed class AuditController(IActionLogService actionLog) : ApiControllerBase
{
    /// <summary>Records that the user opened a screen (the desktop's <c>ActionLogUpdate(12, "frmName")</c> in Form_Load).</summary>
    [HttpPost("screen-opened")]
    public async Task<IActionResult> ScreenOpened(ScreenOpenedRequest request)
    {
        await actionLog.ScreenOpenedAsync(UserCode, request.Screen);
        return NoContent();
    }
}
