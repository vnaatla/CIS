using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Maintenance.BenefitMember;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>Maintenance > Benefit > Member (frmMaintenanceBenefits): Benefit Entry and Benefit Listing tabs (Category = code-tables).</summary>
[Route("api/maintenance/benefit-member")]
[RequireAccess(AccessRight.MaintenanceBenefitMember)]
public sealed class MaintenanceBenefitMemberController(IBenefitMemberService benefits) : ApiControllerBase
{
    /// <summary>Benefit Entry combos and grid rows.</summary>
    [HttpGet("entry/setup")]
    public Task<BenefitEntrySetup> EntrySetup() => benefits.EntrySetupAsync(BenefitScheme.Member);

    /// <summary>Benefit Entry search.</summary>
    [HttpGet("entry")]
    public Task<List<BenefitEntryRow>> EntrySearch([FromQuery] BenefitEntryKey key) => benefits.EntrySearchAsync(BenefitScheme.Member, key);

    /// <summary>Benefit Entry save.</summary>
    [HttpPost("entry")]
    public async Task<object> EntrySave(BenefitEntrySaveRequest request) => new { message = await benefits.EntrySaveAsync(BenefitScheme.Member, request, UserCode) };

    /// <summary>Benefit Listing combos.</summary>
    [HttpGet("listing/setup")]
    public Task<BenefitListingSetup> ListingSetup() => benefits.ListingSetupAsync();

    /// <summary>Benefit Listing search.</summary>
    [HttpPost("listing/search")]
    public Task<List<Dictionary<string, object?>>> Listing(BenefitListingRequest filters) => benefits.ListingAsync(filters);

    /// <summary>Benefit Listing print (Excel of the list).</summary>
    [HttpPost("listing/export")]
    public async Task<FileContentResult> Export(BenefitListingRequest filters)
    {
        var file = await benefits.ListingExportAsync(filters);
        return File(file.Content, file.ContentType, file.FileName);
    }
}
