using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Bordx;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Claim Bordx > Print Bordx (frmPrintBordxExcel and the client variants).</summary>
[Route("api/claim-bordx/print")]
[RequireAccess(AccessRight.BordxPrint)]
public sealed class PrintBordxController(IPrintBordxService print) : ApiControllerBase
{
    /// <summary>Search: the bordereaux and its worksheets.</summary>
    [HttpGet]
    public Task<PrintBordxInfo> Find([FromQuery] string? bordxNo) => print.FindAsync(bordxNo);

    /// <summary>Print / Preview details or summary (Excel); Print stamps the bordereaux as printed.</summary>
    [HttpPost]
    public async Task<FileContentResult> Print(PrintBordxRequest request)
    {
        var file = await print.PrintAsync(request, UserCode);
        return File(file.Content, file.ContentType, file.FileName);
    }
}
