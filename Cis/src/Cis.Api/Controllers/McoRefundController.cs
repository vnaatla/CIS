using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Mco;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>MCO Fees > Refund Listing (frmMCORefundListing).</summary>
[Route("api/mco-fees/refund-listing")]
[RequireAccess(AccessRight.McoFeeListing)]
public sealed class McoRefundController(IMcoRefundService refunds) : ApiControllerBase
{
    /// <summary>Payors, group companies and (for a company) its departments.</summary>
    [HttpGet("lookups")]
    public async Task<object> Lookups([FromQuery] string? company)
    {
        var (payors, companies, departments) = await refunds.LookupsAsync(company);
        return new { payors = payors.Select(p => new { code = p.Code, name = p.Name }), companies, departments };
    }

    /// <summary>Search.</summary>
    [HttpPost("search")]
    public Task<McoRefundResult> Search(McoRefundRequest request) => refunds.SearchAsync(request);

    /// <summary>Print: the listing as Excel.</summary>
    [HttpPost("export")]
    public async Task<FileContentResult> Export(McoRefundRequest request)
    {
        var file = await refunds.ExportAsync(request);
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>Print CR Note (status N numbers the credit note and stamps the members).</summary>
    [HttpPost("credit-note")]
    public async Task<FileContentResult> CreditNote(McoRefundRequest request)
    {
        var file = await refunds.CreditNoteAsync(request, UserCode);
        return File(file.Content, file.ContentType, file.FileName);
    }
}
