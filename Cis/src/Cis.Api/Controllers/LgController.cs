using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Lg;
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
}
