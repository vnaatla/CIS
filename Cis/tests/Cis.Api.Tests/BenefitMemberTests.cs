using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.BenefitMember;
using Cis.Domain.Errors;
using Cis.Infrastructure.Maintenance;

namespace Cis.Api.Tests;

public class BenefitMemberServiceTests
{
    private readonly FakeSessions _sessions = new();
    private readonly FakeMemberRepository _repository = new();
    private readonly BenefitMemberService _service;

    public BenefitMemberServiceTests() =>
        _service = new BenefitMemberService(_sessions, _repository, new NoExporter(), new FixedClock(new DateTime(2026, 11, 1, 10, 0, 0)));

    [Theory]
    [InlineData(null, "H", "A001", "01/01/2004", BenefitMemberService.SearchHealth)]
    [InlineData("N", "", "A001", "01/01/2004", BenefitMemberService.SearchInsured)]
    [InlineData("N", "H", " ", "01/01/2004", BenefitMemberService.SearchPlan)]
    [InlineData("N", "H", "A001", "", BenefitMemberService.SearchDate)]
    [InlineData("N", "H", "A001", "32/01/2004", BenefitMemberService.SearchValidDate)]
    public async Task Entry_search_checks_in_the_desktop_order(string? h, string i, string p, string d, string message) =>
        Assert.Equal(message, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.EntrySearchAsync(BenefitScheme.Member, new(h, i, p, d)))).Message);

    [Theory]
    [InlineData("", "H", "A001", "01/01/2004", BenefitMemberService.SaveHealth)]
    [InlineData("N", null, "A001", "01/01/2004", BenefitMemberService.SaveInsured)]
    [InlineData("N", "H", "", "01/01/2004", BenefitMemberService.SavePlan)]
    [InlineData("N", "H", "A001", "__/__/____", BenefitMemberService.SaveDate)]
    public async Task Entry_save_checks_in_the_desktop_order(string h, string? i, string p, string d, string message) =>
        Assert.Equal(message, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.EntrySaveAsync(BenefitScheme.Member, new(h, i, p, d, []), "DEV"))).Message);

    [Fact]
    public async Task Entry_save_refuses_a_saved_plan_insured_and_date()
    {
        _repository.Existing = true;
        await Assert.ThrowsAsync<ConflictException>(() => _service.EntrySaveAsync(BenefitScheme.Dental, new("N", "H - GROUP FAMILY", "A001", "01/01/2004", []), "DEV"));
    }

    [Fact]
    public async Task Entry_save_writes_only_the_rows_the_desktop_writes()
    {
        List<BenefitEntryRow> rows =
        [
            new("101", "1", "1", "", "", "80"),      // 1 with a percentage → saved
            new("102", "1", "", "", "", ""),         // 1 without RM / % → skipped
            new("103", "3", "", "", "", ""),         // 3 → always saved
            new("104", "4", "", "", "", "50"),       // 4 without RM → skipped
            new("105", "4", "", "12", "50a", ""),    // 4 with RM (any text) → saved, upper-cased
            new("106", "9", "", "", "30", ""),       // not 1–4 → blanked → skipped
            new("107", "", "", "", "30", ""),        // no calculation → skipped
        ];
        Assert.Equal("Record Saved...", await _service.EntrySaveAsync(BenefitScheme.Member, new("n - Non-Sihat", "H - GROUP FAMILY", "A001 ", "01/01/2004", rows), "DEV"));
        Assert.Equal(["A001|H|N|101|1|1|||80", "A001|H|N|103|3||||", "A001|H|N|105|4||12|50A|"], _repository.Inserted);
        Assert.True(_sessions.Last.Committed);
    }

    [Fact]
    public void Listing_filters_are_reduced_like_searchBNF()
    {
        var r = BenefitMemberService.Reduce(new("101 -- CONSULTATION", "S - SIHAT", "H - GROUP FAMILY", "A0012", "x", "AIA - AIA BHD", "B0061", " 1 ", "", null, "", ""));
        Assert.Equal(("101", "S", "H", "A001", "AIA", "B006", "1"), (r.BnfCode, r.HealthType, r.InsuredType, r.PlanCode, r.PayorCode, r.GroupCompany, r.Calculation));
        Assert.Null(r.NoOfDays);
    }

    [Fact]
    public async Task Empty_listing_says_record_not_found_and_print_refuses()
    {
        var f = new BenefitListingRequest(null, null, null, null, null, null, null, null, null, null, null, null);
        Assert.Equal("Record not found", (await Assert.ThrowsAsync<NotFoundException>(() => _service.ListingAsync(f))).Message);
        Assert.Equal("Report cannot be printed without any record.", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.ListingExportAsync(f))).Message);
    }

    private sealed class NoExporter : IReportExporter
    {
        public FileOutput Export(ReportDocument report, ReportFileType type) => new([], "application/octet-stream", "x");
    }

    private sealed class FakeMemberRepository : IBenefitMemberRepository
    {
        public bool Existing { get; set; }
        public List<string> Inserted { get; } = [];

        public Task<(List<LookupOption> Insured, List<LookupOption> Plans, List<LookupOption> Categories)> EntrySetupAsync(IDbSession session, BenefitScheme scheme) =>
            Task.FromResult<(List<LookupOption>, List<LookupOption>, List<LookupOption>)>(([], [], []));

        public Task<BenefitListingSetup> ListingSetupAsync(IDbSession session) => Task.FromResult(new BenefitListingSetup([], [], [], [], [], []));

        public Task<List<BenefitEntryRow>> EntrySearchAsync(IDbSession session, BenefitScheme scheme, string plnCode, string insCode, string hltCode, DateTime effDate) =>
            Task.FromResult(new List<BenefitEntryRow>());

        public Task<bool> EntryExistsAsync(IDbSession session, BenefitScheme scheme, string plnCode, string insCode, DateTime effDate) => Task.FromResult(Existing);

        public Task EntryInsertAsync(IDbSession session, BenefitScheme scheme, string plnCode, string insCode, string hltCode, DateTime effDate, BenefitEntryRow r, string userCode, DateTime today)
        {
            Inserted.Add($"{(scheme == BenefitScheme.Dental ? "D:" : "")}{plnCode}|{insCode}|{hltCode}|{r.BnfCode}|{r.Calculation}|{r.NoOfDays}|{r.NoOfAnn}|{r.RmPerVisit}|{r.Percentage}");
            return Task.CompletedTask;
        }

        public Task<List<Dictionary<string, object?>>> ListingAsync(IDbSession session, BenefitListingRequest filters, DateTime? effDate) =>
            Task.FromResult(new List<Dictionary<string, object?>>());
    }
}

/// <summary>Benefit Member procedures against the real CISDB (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class BenefitMemberDatabaseTests
{
    [Theory]
    [InlineData(BenefitScheme.Member)]
    [InlineData(BenefitScheme.Dental)]
    public async Task Entry_rows_are_saved_found_and_refused_twice(BenefitScheme scheme)
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new BenefitMemberRepository();
        var (insured, plans, categories) = await repo.EntrySetupAsync(s, scheme);
        Assert.NotEmpty(insured);
        Assert.NotEmpty(categories);

        var (plan, date) = (plans[0].Code, new DateTime(1901, 1, 1));
        Assert.False(await repo.EntryExistsAsync(s, scheme, plan, "H", date));
        await repo.EntryInsertAsync(s, scheme, plan, "H", "N", date, new(categories[0].Code, "4", "1", "12", "50", null), "CISTEST", new DateTime(2026, 11, 1));
        Assert.True(await repo.EntryExistsAsync(s, scheme, plan, "H", date));
        var row = Assert.Single(await repo.EntrySearchAsync(s, scheme, plan, "H", "N", date));
        Assert.Equal(("4", "1", "12", "50"), (row.Calculation, row.NoOfDays, row.NoOfAnn, row.RmPerVisit));
        if (scheme != BenefitScheme.Member) return;
        var setup = await repo.ListingSetupAsync(s);
        Assert.NotEmpty(setup.Payors);
        Assert.NotEmpty(setup.GroupCompanies);
        var listed = await repo.ListingAsync(s, new(categories[0].Code, "N", "H", plan, null, null, null, "4", "1", "50", "12", null), date);
        Assert.Contains(listed, r => (DateTime)r["BNFEffDate"]! == date);
    }
}
