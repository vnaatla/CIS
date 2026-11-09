using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.AcExport;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>A/C Export > Invoice Posting > Clinic (GST-2017) (frmPostingInvClnGST2017). The periods are those of api/ac-export/invoice-clinic.</summary>
[Route("api/ac-export/invoice-clinic-gst-2017")]
[RequireAccess(AccessRight.AcExport)]
public sealed class PostInvClnGst2017Controller(IPostInvClnGstService posting, IPostInvClnService clinic) : ApiControllerBase
{
    /// <summary>Search (kind G), Search Debit (D) or Search Credit (C).</summary>
    [HttpPost("search")]
    public Task<PostingResult> Search(PostInvClnRequest request) => posting.SearchAsync(request, true);

    /// <summary>Post: returns the posting workbook; the batch no. is in the X-Batch-No header.</summary>
    [HttpPost("post")]
    public async Task<FileContentResult> Post(PostInvClnPostRequest request)
    {
        var (result, file) = await posting.PostAsync(request, UserCode, true);
        Response.Headers["X-Batch-No"] = result.BatchNo;
        Response.Headers.AccessControlExposeHeaders = "X-Batch-No";
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Print to File: the listed lines as a workbook.</summary>
    [HttpPost("export")]
    public async Task<FileContentResult> Export(PostInvClnRequest request)
    {
        var f = clinic.Export(await posting.SearchAsync(request, true), "PostingInvoiceClinicGST2017");
        return File(f.Content, f.ContentType, f.FileName);
    }

    /// <summary>Generate: the FFS tax invoice of a posted SUN invoice batch.</summary>
    [HttpGet("tax-invoice")]
    public async Task<FileContentResult> TaxInvoice([FromQuery] string? invBatch)
    {
        var f = await posting.TaxInvoiceAsync(invBatch);
        return File(f.Content, f.ContentType, f.FileName);
    }
}
