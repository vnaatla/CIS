using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Lg;
using Cis.Domain.Errors;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>LG &amp; Audit > Generate LG (frmLGGeneration).</summary>
[Route("api/lg-audit/generate-lg")]
[RequireAccess(AccessRight.GenerateLg)]
public sealed class LgController(ILgService lg) : ApiControllerBase
{
    /// <summary>Drop-downs (GL types, decline reasons, delay reasons).</summary>
    [HttpGet("lookups")]
    public Task<LgLookups> Lookups() => lg.LookupsAsync();

    /// <summary>Search: the member's covers.</summary>
    [HttpGet("members")]
    public Task<LgSearchResult> Search([FromQuery] string? member) => lg.SearchAsync(member);

    /// <summary>A cover: limits, LG checks and its LGs.</summary>
    [HttpGet("members/{member}/{cover}")]
    public Task<LgMemberInfo> Member(string member, string cover) => lg.MemberAsync(member, cover);

    /// <summary>GP / specialist cases of the member (cover optional).</summary>
    [HttpGet("cases")]
    public Task<List<Dictionary<string, object?>>> Cases([FromQuery] string? member, [FromQuery] string? cover, [FromQuery] bool specialist) =>
        lg.CasesAsync(member, cover, specialist);

    /// <summary>The group company's clinics.</summary>
    [HttpGet("clinics")]
    public async Task<object> Clinics([FromQuery] string? company) => (await lg.ClinicsAsync(company)).Select(c => new { code = c.Code, name = c.Name });

    /// <summary>One LG.</summary>
    [HttpGet("{lgNo}")]
    public Task<LgDetail> Load(string lgNo, [FromQuery] string? groupCompany, [FromQuery] string? category) => lg.LoadAsync(lgNo, groupCompany, category);

    /// <summary>Add: the limit checks for the cover (confirmed = OK answers so far).</summary>
    [HttpGet("members/{member}/{cover}/add-check")]
    public Task<LgGateResult> AddCheck(string member, string cover, [FromQuery] int confirmed) => lg.AddCheckAsync(member, cover, confirmed);

    /// <summary>Edit: the case, clinic, expiry and limit checks of the LG.</summary>
    [HttpGet("{lgNo}/edit-check")]
    public Task<LgGateResult> EditCheck(string lgNo, [FromQuery] string? clinic, [FromQuery] string? category, [FromQuery] int confirmed) =>
        lg.EditCheckAsync(lgNo, clinic, category, confirmed);

    /// <summary>LG / Update: issues the LG (questions come back as Confirm until answered).</summary>
    [HttpPost("issue")]
    public Task<LgActionResult> Issue(LgIssueRequest request) => lg.IssueAsync(request, UserCode);

    /// <summary>Decline.</summary>
    [HttpPost("decline")]
    public Task<LgActionResult> Decline(LgDeclineRequest request) => lg.DeclineAsync(request, UserCode);

    /// <summary>Preview: the GL letter as it stands.</summary>
    [HttpPost("preview")]
    public async Task<FileContentResult> Preview(LgIssueRequest request)
    {
        var file = await lg.PreviewAsync(request, UserCode);
        return File(file.Content, file.ContentType, file.FileName);
    }

    /// <summary>A document of the LG.</summary>
    [HttpGet("{lgNo}/documents/file")]
    public async Task<FileContentResult> Document(string lgNo, [FromQuery] string? name)
    {
        var (content, fileName) = await lg.DocumentAsync(lgNo, name);
        return File(content, "application/octet-stream", fileName);
    }

    /// <summary>Copy File: puts a file into the LG's folder.</summary>
    [HttpPost("{lgNo}/documents")]
    [RequestSizeLimit(50_000_000)]
    public async Task<object> Copy(string lgNo, IFormFile? file)
    {
        if (file is null || file.Length == 0) throw new BusinessRuleException("Please choose a file.");
        await using var stream = file.OpenReadStream();
        return new { message = await lg.CopyAsync(lgNo, file.FileName, stream) };
    }
}
