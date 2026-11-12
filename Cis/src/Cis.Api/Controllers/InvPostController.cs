using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.AcExport;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>A/C Export > Invoice Posting (frmInvPosting). The payors are those of api/payment/post/clinic/lookups.</summary>
[Route("api/ac-export/invoice-posting")]
[RequireAccess(AccessRight.AcExport)]
public sealed class InvPostController(IInvPostService posting, IPostInvClnService clinic) : ApiControllerBase
{
    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<PostingResult> Search(InvPostRequest request) => posting.SearchAsync(request);

    /// <summary>Post: the posting workbook; the listed cases are marked posted.</summary>
    [HttpPost("post")]
    public async Task<FileContentResult> Post(InvPostRequest request)
    {
        var f = await posting.PostAsync(request, UserCode);
        return File(f.Content, f.ContentType, f.FileName);
    }

    /// <summary>Print to File: the listed lines as a workbook.</summary>
    [HttpPost("export")]
    public async Task<FileContentResult> Export(InvPostRequest request)
    {
        var f = clinic.Export(await posting.SearchAsync(request), "InvoicePostingList");
        return File(f.Content, f.ContentType, f.FileName);
    }
}
