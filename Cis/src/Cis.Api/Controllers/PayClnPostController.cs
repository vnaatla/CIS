using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Payment;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Payment > Pay Clinic Post (frmPayClinicPost).</summary>
[Route("api/payment/post/clinic")]
[RequireAccess(AccessRight.AcExport)]
public sealed class PayClnPostController(IPayClnPostService post, Application.Bordx.IBordxInvService lookups) : ApiControllerBase
{
    /// <summary>The payors (cboPAYCode).</summary>
    [HttpGet("lookups")]
    public Task<object> Lookups() => lookups.LookupsAsync();

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<PayClnPostResult> Search(PayClnPostRequest request) => post.SearchAsync(request);

    /// <summary>Post: the posting workbook; the listed cases are marked posted.</summary>
    [HttpPost]
    public async Task<FileContentResult> Post(PayClnPostRequest request)
    {
        var f = await post.PostAsync(request, UserCode);
        return File(f.Content, f.ContentType, f.FileName);
    }
}
