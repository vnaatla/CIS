using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Membership;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.StaticFiles;

namespace Cis.Api.Controllers;

/// <summary>Invoice > General Exclusions (Enquiry) (frmGeneralBenefitExcEnq), read-only.</summary>
[Route("api/invoice/general-exclusions")]
[RequireAccess(AccessRight.GeneralExclusionsEnquiry)]
public sealed class GeneralExclusionsController(IGeneralExclusionsService exclusions) : ApiControllerBase
{
    private static readonly FileExtensionContentTypeProvider ContentTypes = new();

    /// <summary>Group company names.</summary>
    [HttpGet("group-companies")]
    public Task<List<string>> GroupCompanies() => exclusions.GroupCompaniesAsync();

    /// <summary>Search: the company's exclusions and remarks.</summary>
    [HttpGet]
    public Task<GeneralExclusionsView> Search([FromQuery] string? groupCompany) => exclusions.SearchAsync(groupCompany);

    /// <summary>Policy Wording tab: file names (configured = false when the folder setting is missing).</summary>
    [HttpGet("policy-wording")]
    public object PolicyFiles([FromQuery] string? groupCompany) =>
        exclusions.PolicyFiles(groupCompany) is { } files ? new { configured = true, files } : new { configured = false, files = new List<string>() };

    /// <summary>One policy wording file.</summary>
    [HttpGet("policy-wording/file")]
    public async Task<IActionResult> PolicyFile([FromQuery] string? groupCompany, [FromQuery] string? fileName)
    {
        var (content, name) = await exclusions.PolicyFileAsync(groupCompany, fileName);
        return File(content, ContentTypes.TryGetContentType(name, out var type) ? type : "application/octet-stream", name);
    }
}
