using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Clinic > Clinic (frmInvoice).</summary>
[Route("api/invoice/clinic")]
[RequireAccess(AccessRight.InvoiceClinic)]
public sealed class ClinicInvoiceController(IClinicInvoiceService invoices) : ApiControllerBase
{
    /// <summary>Receipt modes and process months.</summary>
    [HttpGet("setup")]
    public Task<ClinicSetup> Setup() => invoices.SetupAsync();

    /// <summary>A panel clinic.</summary>
    [HttpGet("clinic/{code}")]
    public Task<ClinicInfo> Clinic(string code) => invoices.ClinicAsync(code);

    /// <summary>A diagnosis (treatment reason).</summary>
    [HttpGet("diagnosis/{code}")]
    public Task<InvoiceDiagnosis> Diagnosis(string code) => invoices.DiagnosisAsync(code);

    /// <summary>Go: the member panel.</summary>
    [HttpGet("member/{mbmNumber}")]
    public Task<InvoiceMember> Member(string mbmNumber, [FromQuery] string? covId) => invoices.MemberAsync(mbmNumber, covId);

    /// <summary>Add: the line and its worksheet (or the duplicate question).</summary>
    [HttpPost("add")]
    public Task<ClinicAddResult> Add(ClinicInvoiceRequest request) => invoices.AddAsync(request, UserCode);

    /// <summary>Verify: the errors (empty = verified).</summary>
    [HttpPost("verify")]
    public async Task<object> Verify(ClinicInvoiceRequest request)
    {
        var errors = await invoices.VerifyAsync(request, UserCode);
        return new { errors, message = errors.Count == 0 ? ClinicInvoiceService.Verified : null };
    }

    /// <summary>Save: the new case number.</summary>
    [HttpPost]
    public async Task<object> Save(ClinicInvoiceRequest request) => new { caseNo = await invoices.SaveAsync(request, UserCode), message = ClinicInvoiceService.Saved };
}
