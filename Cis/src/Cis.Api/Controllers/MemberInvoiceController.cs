using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Invoice;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Invoice > Clinic > Member (frmInvoiceMBM). Set-up, diagnosis and Go reuse <c>api/invoice/clinic</c>'s logic under this right.</summary>
[Route("api/invoice/member")]
[RequireAccess(AccessRight.InvoiceMember)]
public sealed class MemberInvoiceController(IMemberInvoiceService invoices, IClinicInvoiceService clinic) : ApiControllerBase
{
    /// <summary>Receipt modes.</summary>
    [HttpGet("setup")]
    public Task<ClinicSetup> Setup() => clinic.SetupAsync();

    /// <summary>A diagnosis.</summary>
    [HttpGet("diagnosis/{code}")]
    public Task<InvoiceDiagnosis> Diagnosis(string code) => clinic.DiagnosisAsync(code);

    /// <summary>Go: the member panel.</summary>
    [HttpGet("member/{mbmNumber}")]
    public Task<InvoiceMember> Member(string mbmNumber, [FromQuery] string? covId) => clinic.MemberAsync(mbmNumber, covId);

    /// <summary>A clinic (NP / NPS take a typed clinic name).</summary>
    [HttpGet("clinic/{code}")]
    public Task<ClinicInfo> Clinic(string code) => invoices.ClinicAsync(code);

    /// <summary>Add: the line and its worksheet.</summary>
    [HttpPost("add")]
    public Task<MemberInvoiceLine> Add(MemberInvoiceRequest request) => invoices.AddAsync(request, UserCode);

    /// <summary>Verify.</summary>
    [HttpPost("verify")]
    public Task<MemberVerifyResult> Verify(MemberInvoiceRequest request) => invoices.VerifyAsync(request, UserCode);

    /// <summary>Save: the new case number.</summary>
    [HttpPost]
    public async Task<object> Save(MemberInvoiceRequest request) => new { caseNo = await invoices.SaveAsync(request, UserCode), message = ClinicInvoiceService.Saved };
}
