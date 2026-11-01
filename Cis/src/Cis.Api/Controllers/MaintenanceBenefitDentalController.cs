using Cis.Api.Auth;
using Cis.Api.Common;
using Cis.Application.Maintenance.BenefitMember;
using Cis.Domain.Security;
using Microsoft.AspNetCore.Mvc;

namespace Cis.Api.Controllers;

/// <summary>
/// Maintenance > Benefit > Dental (frmMaintenanceBNFDental) tab Benefit Entry: DentalBNF over DentalBNFCategory, the Member tab's rules.
/// Tooth Category / Tooth Reference / Benefit Category are code-tables. MDIMain opens the form with the Benefit Member right.
/// </summary>
[Route("api/maintenance/benefit-dental")]
[RequireAccess(AccessRight.MaintenanceBenefitMember)]
public sealed class MaintenanceBenefitDentalController(IBenefitMemberService benefits) : ApiControllerBase
{
    /// <summary>Benefit Entry combos and grid rows.</summary>
    [HttpGet("entry/setup")]
    public Task<BenefitEntrySetup> EntrySetup() => benefits.EntrySetupAsync(BenefitScheme.Dental);

    /// <summary>Benefit Entry search.</summary>
    [HttpGet("entry")]
    public Task<List<BenefitEntryRow>> EntrySearch([FromQuery] BenefitEntryKey key) => benefits.EntrySearchAsync(BenefitScheme.Dental, key);

    /// <summary>Benefit Entry save.</summary>
    [HttpPost("entry")]
    public async Task<object> EntrySave(BenefitEntrySaveRequest request) =>
        new { message = await benefits.EntrySaveAsync(BenefitScheme.Dental, request, UserCode) };
}
