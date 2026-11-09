using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.AcExport;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>A/C Export > Invoice Posting > Clinic (frmPostingInvoiceCln).</summary>
[Route("api/ac-export/invoice-clinic")]
[RequireAccess(AccessRight.AcExport)]
public sealed class PostInvClnController(IPostInvClnService posting) : ApiControllerBase
{
    /// <summary>The accounting periods.</summary>
    [HttpGet("periods")]
    public Task<List<string>> Periods() => posting.PeriodsAsync();

    /// <summary>Search (kind G), Search Debit (D) or Search Credit (C).</summary>
    [HttpPost("search")]
    public Task<PostingResult> Search(PostInvClnRequest request) => posting.SearchAsync(request);

    /// <summary>Post: returns the posting workbook; the batch no. is in the X-Batch-No header.</summary>
    [HttpPost("post")]
    public async Task<FileContentResult> Post(PostInvClnPostRequest request)
    {
        var (result, file) = await posting.PostAsync(request, UserCode);
        Response.Headers["X-Batch-No"] = result.BatchNo;
        Response.Headers.AccessControlExposeHeaders = "X-Batch-No";
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Print to File: the listed lines as a workbook.</summary>
    [HttpPost("export")]
    public async Task<FileContentResult> Export(PostInvClnRequest request)
    {
        var f = posting.Export(await posting.SearchAsync(request), "PostingInvoiceClinic");
        return File(f.Content, f.ContentType, f.FileName);
    }
}
