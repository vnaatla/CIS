using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Email Notification (Payment) (frmEmailNotification).</summary>
[Route("api/email/payment")]
[RequireAccess(AccessRight.EmailNotification)]
public sealed class EmlPayController(IEmlPayService email) : ApiControllerBase
{
    /// <summary>Group companies.</summary>
    [HttpGet("companies")]
    public Task<List<string>> Companies() => email.CompaniesAsync();

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(EmlPayRequest request) => email.SearchAsync(request);

    /// <summary>Send the e-mails of the ticked PVs and mark them as e-mailed.</summary>
    [HttpPost("send")]
    public async Task<object> Send(EmlPaySendRequest request)
    {
        var sent = await email.SendAsync(request);
        return new { sent, message = EmlPayService.Sent };
    }
}
