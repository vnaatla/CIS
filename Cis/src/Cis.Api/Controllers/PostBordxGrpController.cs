using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.AcExport;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>A/C Export > Bordx Posting > Grp Company: the actions shared by the two screens. The periods are those of api/ac-export/invoice-clinic, the group
/// companies those of api/ac-export/invoice-member.</summary>
[RequireAccess(AccessRight.AcExport)]
public abstract class PostBordxGrpControllerBase(IPostBordxGrpService posting, IPostInvClnService clinic, bool gst) : ApiControllerBase
{
    /// <summary>Search (kind G), Search Debit (D) or Search Credit (C).</summary>
    [HttpPost("search")]
    public Task<PostingResult> Search(PostBordxGrpRequest request) => posting.SearchAsync(request, gst);

    /// <summary>Post: returns the posting workbook; the batch no. is in the X-Batch-No header.</summary>
    [HttpPost("post")]
    public async Task<FileContentResult> Post(PostBordxGrpPostRequest request)
    {
        var (result, file) = await posting.PostAsync(request, UserCode, gst);
        Response.Headers["X-Batch-No"] = result.BatchNo;
        Response.Headers.AccessControlExposeHeaders = "X-Batch-No";
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Print to File: the listed lines as a workbook.</summary>
    [HttpPost("export")]
    public async Task<FileContentResult> Export(PostBordxGrpRequest request)
    {
        var f = clinic.Export(await posting.SearchAsync(request, gst), "PostingBordxGroup");
        return File(f.Content, f.ContentType, f.FileName);
    }
}

/// <summary>A/C Export > Bordx Posting > Grp Company (frmPostingBordxtoGrpComp, hidden in the desktop menu).</summary>
[Route("api/ac-export/bordx-group")]
public sealed class PostBordxGrpController(IPostBordxGrpService posting, IPostInvClnService clinic) : PostBordxGrpControllerBase(posting, clinic, false);

/// <summary>A/C Export > Bordx Posting > Grp Company (GST) (frmPostingBordxtoGrpGST).</summary>
[Route("api/ac-export/bordx-group-gst")]
public sealed class PostBordxGrpGstController(IPostBordxGrpService posting, IPostInvClnService clinic) : PostBordxGrpControllerBase(posting, clinic, true);
