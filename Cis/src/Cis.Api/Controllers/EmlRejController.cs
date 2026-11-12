using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Reports;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Email Notification (Rejection) (frmEmailRejectedClaim). The group companies are those of api/email/payment/companies.</summary>
[Route("api/email/rejection")]
[RequireAccess(AccessRight.EmailNotification)]
public sealed class EmlRejController(IEmlRejService email) : ApiControllerBase
{
    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<List<Dictionary<string, object?>>> Search(EmlRejRequest request) => email.SearchAsync(request);

    /// <summary>Send the e-mails (BSN members without a personal e-mail get a letter in the returned zip) and mark the claims as notified.</summary>
    [HttpPost("send")]
    public async Task<object> Send(EmlRejSendRequest request)
    {
        var r = await email.SendAsync(request);
        return new { r.Sent, r.Letters, r.LettersFileName, r.LettersZip, message = EmlPayService.Sent };
    }
}
