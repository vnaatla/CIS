using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.AcExport;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>A/C Export > Payment Posting > Clinic (FrmPostingPay2Cln). The periods are those of api/ac-export/invoice-clinic.</summary>
[Route("api/ac-export/payment-clinic")]
[RequireAccess(AccessRight.AcExport)]
public sealed class PostPayClnController(IPostPayClnService posting, IPostInvClnService clinic) : ApiControllerBase
{
    /// <summary>Search (kind G), Search Debit (D) or Search Credit (C).</summary>
    [HttpPost("search")]
    public Task<PostingResult> Search(PostPayClnRequest request) => posting.SearchAsync(request);

    /// <summary>Post: returns the posting workbook; the batch no. is in the X-Batch-No header.</summary>
    [HttpPost("post")]
    public async Task<FileContentResult> Post(PostPayClnPostRequest request)
    {
        var (result, file) = await posting.PostAsync(request, UserCode);
        Response.Headers["X-Batch-No"] = result.BatchNo;
        Response.Headers.AccessControlExposeHeaders = "X-Batch-No";
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Print to File: the listed lines as a workbook.</summary>
    [HttpPost("export")]
    public async Task<FileContentResult> Export(PostPayClnRequest request)
    {
        var f = clinic.Export(await posting.SearchAsync(request), "PostingPaymentClinic");
        return File(f.Content, f.ContentType, f.FileName);
    }
}
