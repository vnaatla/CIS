using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Clinic > Enquiry (FrmInvoiceEnq).</summary>
[Route("api/invoice/enquiry")]
[RequireAccess(AccessRight.InvoiceAdjustment)]
public sealed class InvoiceEnquiryController(IInvoiceEnquiryService enquiry) : ApiControllerBase
{
    /// <summary>The case, its member, stored worksheet and bordereaux.</summary>
    [HttpGet("{casNumber}")]
    public Task<InvoiceEnquiryCase> Find(string casNumber) => enquiry.FindAsync(casNumber);

    /// <summary>The next / previous case number.</summary>
    [HttpGet("{casNumber}/step")]
    public async Task<object> Step(string casNumber, [FromQuery] bool next) => new { casNumber = await enquiry.StepAsync(casNumber, next) };

    /// <summary>Payments and receipts.</summary>
    [HttpGet("{casNumber}/payments")]
    public async Task<object> Payments(string casNumber)
    {
        var (payments, receipts) = await enquiry.PaymentsAsync(casNumber);
        return new { payments, receipts };
    }

    /// <summary>Medication requests.</summary>
    [HttpGet("{casNumber}/medication")]
    public async Task<object> Medication(string casNumber)
    {
        var (requests, medicines) = await enquiry.MedRequestsAsync(casNumber);
        return new { requests, medicines };
    }

    /// <summary>The case's documents.</summary>
    [HttpGet("{casNumber}/documents")]
    public List<CaseDocument> Documents(string casNumber) =>
        enquiry.Documents(casNumber) ?? throw new NotConfiguredException("The outpatient documents folder is not configured (Files:Folders:Outpatient).");

    /// <summary>A document.</summary>
    [HttpGet("{casNumber}/documents/file")]
    public async Task<FileContentResult> Document(string casNumber, [FromQuery] string? group, [FromQuery] string? name)
    {
        var (content, fileName) = await enquiry.DocumentAsync(casNumber, group, name);
        return File(content, "application/octet-stream", fileName);
    }
}
