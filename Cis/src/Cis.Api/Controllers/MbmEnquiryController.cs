using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Membership;
using Cis.Domain.Errors;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Membership > MBM Enquiry (FrmMBMCLNEnq). Prev / Next and the history tabs share the MBM Adjustment logic.</summary>
[Route("api/membership/enquiry")]
[RequireAccess(AccessRight.MbmEnquiry)]
public sealed class MbmEnquiryController(IMbmEnquiryService enquiry, IMbmAdjustmentService adjustment, IGeneralExclusionsService exclusions) : ApiControllerBase
{
    /// <summary>Show: the member with its limits and covers.</summary>
    [HttpGet]
    public Task<EnqMember> Find([FromQuery] string? member) => enquiry.FindAsync(member);

    /// <summary>Prev / Next member number.</summary>
    [HttpGet("step")]
    public async Task<object> Step([FromQuery] string? member, [FromQuery] bool next) => new { member = await adjustment.StepAsync(member, next) };

    /// <summary>Adjustment history tab.</summary>
    [HttpGet("amend-history")]
    public Task<List<Dictionary<string, object?>>> AmendHistory([FromQuery] string? member) => adjustment.AmendHistoryAsync(member);

    /// <summary>MBM history tab (by IC).</summary>
    [HttpGet("mbm-history")]
    public Task<List<Dictionary<string, object?>>> MbmHistory([FromQuery] string? ic) => adjustment.MbmHistoryAsync(ic);

    /// <summary>FlexCare (cover 00) / Family Benefits (other covers) tab.</summary>
    [HttpGet("flex")]
    public Task<EnqFlex> Flex([FromQuery] string? member, [FromQuery] string? coverId) => enquiry.FlexAsync(member, coverId);

    /// <summary>Case tab.</summary>
    [HttpGet("cases")]
    public Task<EnqCases> Cases([FromQuery] string? member) => enquiry.CasesAsync(member);

    /// <summary>Benefit Schedule tab: the group company's policy wording files.</summary>
    [HttpGet("policy-files")]
    public object PolicyFiles([FromQuery] string? groupCompany)
    {
        var files = enquiry.PolicyFiles(groupCompany);
        return new { configured = files is not null, files = files ?? [] };
    }

    /// <summary>Open a listed policy wording file.</summary>
    [HttpGet("policy-file")]
    public async Task<IActionResult> PolicyFile([FromQuery] string? groupCompany, [FromQuery] string? fileName)
    {
        if (enquiry.PolicyFiles(groupCompany)?.Contains(fileName?.Trim() ?? "", StringComparer.OrdinalIgnoreCase) != true) throw new NotFoundException("File not found.");
        var (content, name) = await exclusions.PolicyFileAsync(groupCompany, fileName);
        return File(content, name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? "application/pdf" : "image/jpeg", name);
    }
}
