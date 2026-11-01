using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.BenefitMember;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IBenefitMemberRepository"/> on the <c>cis.BenefitMember_*</c> / <c>cis.BenefitDental_*</c> procedures.</summary>
public sealed class BenefitMemberRepository : IBenefitMemberRepository
{
    /// <inheritdoc />
    public async Task<(List<LookupOption> Insured, List<LookupOption> Plans, List<LookupOption> Categories)> EntrySetupAsync(IDbSession session, BenefitScheme scheme)
    {
        using var grid = await session.MultipleAsync(Proc(scheme, "EntrySetup"));
        return ((await grid.ReadAsync<LookupOption>()).ToList(), (await grid.ReadAsync<LookupOption>()).ToList(), (await grid.ReadAsync<LookupOption>()).ToList());
    }

    /// <inheritdoc />
    public async Task<BenefitListingSetup> ListingSetupAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.BenefitMember_ListingSetup");
        async Task<List<LookupOption>> Next() => (await grid.ReadAsync<LookupOption>()).ToList();
        return new BenefitListingSetup(await Next(), await Next(), await Next(), await Next(), await Next(), await Next());
    }

    /// <inheritdoc />
    public Task<List<BenefitEntryRow>> EntrySearchAsync(IDbSession session, BenefitScheme scheme, string plnCode, string insCode, string hltCode, DateTime effDate) =>
        session.QueryAsync<BenefitEntryRow>(Proc(scheme, "EntrySearch"), new { plnCode, insCode, hltCode, effDate });

    /// <inheritdoc />
    public Task<bool> EntryExistsAsync(IDbSession session, BenefitScheme scheme, string plnCode, string insCode, DateTime effDate) =>
        session.ScalarAsync<bool>(Proc(scheme, "EntryExists"), new { plnCode, insCode, effDate });

    /// <inheritdoc />
    public Task EntryInsertAsync(IDbSession session, BenefitScheme scheme, string plnCode, string insCode, string hltCode, DateTime effDate, BenefitEntryRow row, string userCode,
        DateTime today) =>
        session.ExecAsync(Proc(scheme, "EntryInsert"), new
        {
            plnCode, insCode, hltCode, effDate, bnfCode = row.BnfCode, calculation = row.Calculation, noOfDays = SmallInt(row.NoOfDays),
            noOfAnn = SmallInt(row.NoOfAnn), rmPerVisit = row.RmPerVisit, percentage = row.Percentage, user = userCode, today,
        });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> ListingAsync(IDbSession session, BenefitListingRequest f, DateTime? effDate) =>
        session.RowsAsync("cis.BenefitMember_Listing", new
        {
            bnfCode = f.BnfCode, hltCode = f.HealthType, insCode = f.InsuredType, plnCode = f.PlanCode, effDate, payCode = f.PayorCode,
            grpCompany = f.GroupCompany, calculation = f.Calculation, noOfDays = f.NoOfDays, rmPerVisit = f.RmPerVisit, noOfAnn = f.NoOfAnn,
            percentage = f.Percentage,
        });

    /// <summary>cis.BenefitMember_* or cis.BenefitDental_* (fixed names, no user input).</summary>
    private static string Proc(BenefitScheme scheme, string name) => (scheme == BenefitScheme.Dental ? "cis.BenefitDental_" : "cis.BenefitMember_") + name;

    private static short? SmallInt(string? v) => short.TryParse(v, out var n) ? n : null;
}
