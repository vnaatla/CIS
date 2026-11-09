using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Receipt;
using Cis.Domain.Errors;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Receipt > Upload Receipt (frmReceiptUpload) and Upload Receipt (Multiple Receipts) (frmReceiptUpMultiple).</summary>
[Route("api/receipt/upload")]
[RequireAccess(AccessRight.ReceiptPayor)]
public sealed class RcpUpController(IRcpUpService upload) : ApiControllerBase
{
    /// <summary>Upload the spreadsheet: its rows compared with CIS.</summary>
    [HttpPost]
    [RequestSizeLimit(20_000_000)]
    public async Task<List<RcpUpRow>> Upload(IFormFile? file, [FromQuery] bool multiple = false)
    {
        if (file is null || file.Length == 0) throw new BusinessRuleException(RcpUpService.NoFile);
        await using var copy = new MemoryStream();
        await file.CopyToAsync(copy);
        copy.Position = 0;
        return await upload.UploadAsync(copy, multiple);
    }

    /// <summary>Update: the matched rows' receipts.</summary>
    [HttpPost("save")]
    public async Task<object> Save(RcpUpSaveRequest request) => new { message = await upload.SaveAsync(request, UserCode) };
}
