using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Email Notification (Providers) (frmEmailNotificationProvider).</summary>
[Route("api/email/providers")]
[RequireAccess(AccessRight.EmailNotification)]
public sealed class EmlPrvController(IEmlPrvService email) : ApiControllerBase
{
    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(EmlPrvRequest request) => email.SearchAsync(request);

    /// <summary>Send the e-mails of the ticked clinic PVs and mark them as e-mailed.</summary>
    [HttpPost("send")]
    public async Task<object> Send(EmlPrvSendRequest request)
    {
        var sent = await email.SendAsync(request);
        return new { sent, message = EmlPayService.Sent };
    }
}
