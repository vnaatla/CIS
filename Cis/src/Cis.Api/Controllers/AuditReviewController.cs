using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Audit;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>LG &amp; Audit > Audit review (frmAuditInvoice).</summary>
[Route("api/lg-audit/audit-review")]
[RequireAccess(AccessRight.AuditReview)]
public sealed class AuditReviewController(IAuditReviewService audit) : ApiControllerBase
{
    /// <summary>Go: the case and its audit details.</summary>
    [HttpGet("{casNumber}")]
    public Task<AuditCase> Load(string casNumber) => audit.LoadAsync(casNumber);

    /// <summary>Update the audit details.</summary>
    [HttpPost("{casNumber}")]
    public async Task<object> Update(string casNumber, AuditReviewRequest request) => new { message = await audit.UpdateAsync(casNumber, request) };

    /// <summary>The next / previous case number.</summary>
    [HttpGet("{casNumber}/step")]
    public async Task<object> Step(string casNumber, [FromQuery] bool next) => new { casNumber = await audit.StepAsync(casNumber, next) };

    /// <summary>The case's guarantee letter documents.</summary>
    [HttpGet("{casNumber}/documents")]
    public List<string> Documents(string casNumber) =>
        audit.Documents(casNumber) ?? throw new NotConfiguredException("The outpatient documents folder is not configured (Files:Folders:Outpatient).");

    /// <summary>A document.</summary>
    [HttpGet("{casNumber}/documents/file")]
    public async Task<FileContentResult> Document(string casNumber, [FromQuery] string? name)
    {
        var (content, fileName) = await audit.DocumentAsync(casNumber, name);
        return File(content, "application/octet-stream", fileName);
    }

    /// <summary>Copy a file into the case's guarantee letter folder.</summary>
    [HttpPost("{casNumber}/documents")]
    [RequestSizeLimit(50_000_000)]
    public async Task<object> Upload(string casNumber, IFormFile? file)
    {
        if (file is null || file.Length == 0) throw new BusinessRuleException("Please choose a file.");
        await using var stream = file.OpenReadStream();
        return new { message = await audit.UploadAsync(casNumber, file.FileName, stream) };
    }
}
