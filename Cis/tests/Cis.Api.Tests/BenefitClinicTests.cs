using Cis.Application.Abstractions;
using Cis.Application.Lookups;
using Cis.Application.Maintenance.BenefitClinic;
using Cis.Application.Maintenance.LgTemplates;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Maintenance;

namespace Cis.Api.Tests;

public class BenefitClinicServiceTests
{
    private readonly FakeSessions _sessions = new();
    private readonly FakeClinicRepository _repository = new();
    private readonly BenefitClinicService _service;

    public BenefitClinicServiceTests() =>
        _service = new BenefitClinicService(_sessions, _repository, new FixedClock(new DateTime(2026, 11, 1, 10, 0, 0)));

    [Theory]
    [InlineData(null, "01/01/2002", BenefitClinicService.SearchType)]
    [InlineData("Gold", "", BenefitClinicService.SearchDate)]
    [InlineData("Gold", "31/02/2002", BenefitClinicService.SearchValidDate)]
    public async Task Search_checks_in_the_desktop_order(string? type, string date, string message) =>
        Assert.Equal(message, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SearchAsync(type, date))).Message);

    [Fact]
    public async Task Search_without_rows_says_record_not_found() =>
        Assert.Equal("Record not found!", (await Assert.ThrowsAsync<NotFoundException>(() => _service.SearchAsync("Gold", "01/01/2099"))).Message);

    [Theory]
    [InlineData(" ", "01/01/2002", BenefitClinicService.SaveType)]
    [InlineData("Gold", "__/__/____", BenefitClinicService.SaveDate)]
    public async Task Save_checks_type_then_date(string type, string date, string message) =>
        Assert.Equal(message, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(new(type, date, null), "DEV"))).Message);

    [Fact]
    public async Task Save_refuses_a_saved_type_and_date()
    {
        _repository.Existing = true;
        var ex = await Assert.ThrowsAsync<ConflictException>(() => _service.SaveAsync(new("Gold", "01/01/2002", null), "DEV"));
        Assert.Equal("Record Duplicated! Please Change The Corresponding Field", ex.Message);
        Assert.Empty(_repository.Inserted);
    }

    [Fact]
    public async Task Save_writes_one_row_per_category_in_one_transaction()
    {
        var message = await _service.SaveAsync(new(" Gold ", "05/03/2026", new() { ["A1"] = " 12.5 ", ["B2"] = "abc" }), "DEV");
        Assert.Equal("Record Saved...", message);
        Assert.Equal(["A1|12.5|Gold|2026-03-05|DEV|2026-11-01", "B2||Gold|2026-03-05|DEV|2026-11-01", "C3||Gold|2026-03-05|DEV|2026-11-01"], _repository.Inserted);
        Assert.True(_sessions.Last.Committed);
    }

    private sealed class FakeClinicRepository : IBenefitClinicRepository
    {
        public bool Existing { get; set; }
        public List<string> Inserted { get; } = [];

        public Task<List<LookupOption>> CategoriesAsync(IDbSession session) => Task.FromResult<List<LookupOption>>([new("A1", "A"), new("B2", "B"), new("C3", "C")]);

        public Task<List<ClinicBenefitFee>> SearchAsync(IDbSession session, string cardType, DateTime effDate) => Task.FromResult(new List<ClinicBenefitFee>());

        public Task<bool> ExistsAsync(IDbSession session, string cardType, DateTime effDate) => Task.FromResult(Existing);

        public Task InsertAsync(IDbSession session, string cardType, DateTime effDate, string bnfCode, string? rmPerVisit, string userCode, DateTime today)
        {
            Inserted.Add($"{bnfCode}|{rmPerVisit}|{cardType}|{effDate:yyyy-MM-dd}|{userCode}|{today:yyyy-MM-dd}");
            return Task.CompletedTask;
        }
    }
}

public class LgBindingServiceTests
{
    [Fact]
    public async Task Search_needs_a_group_company()
    {
        var service = new LgBindingService(new FakeSessions(), null!);
        Assert.Equal("Plese Select Group Company to Search", (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(" ", null, null))).Message);
    }

    [Theory]
    [InlineData("", "N")]
    [InlineData("HCDependent.dot", "Y")]
    public void Dependent_indicator_follows_the_dependent_template(string template, string expected) =>
        Assert.Equal(expected, LgBindingService.DepInd(new(null, null, null, null, null, null, null, template, null)));
}

/// <summary>Benefit Clinic and LG binding procedures against the real CISDB (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class BenefitClinicDatabaseTests
{
    [Fact]
    public async Task Clinic_benefit_rows_are_saved_found_and_refused_twice()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new BenefitClinicRepository();
        var date = new DateTime(1901, 1, 1);
        var categories = await repo.CategoriesAsync(s);
        Assert.NotEmpty(categories);
        Assert.False(await repo.ExistsAsync(s, "ZZTEST", date));
        await repo.InsertAsync(s, "ZZTEST", date, categories[0].Code, "9.5", "CISTEST", new DateTime(2026, 11, 1));
        Assert.True(await repo.ExistsAsync(s, "ZZTEST", date));
        var row = Assert.Single(await repo.SearchAsync(s, "ZZTEST", date));
        Assert.Equal((categories[0].Code, "9.5"), (row.BnfCode, row.RmPerVisit));
    }

    [Fact]
    public async Task Lg_binding_lookups_and_search_run()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new LgBindingRepository();
        var lookups = await repo.LookupsAsync(s);
        Assert.NotEmpty(lookups.GroupCompanies);
        await repo.InsertAsync(s, new("ZZ CIS WEB TEST", "ZZ1", "Normal GL", "T.dot", "ZZ", "H", "N", "", ""), "N");
        var row = Assert.Single(await repo.SearchAsync(s, "%ZZ CIS WEB TEST%", null, "ZZ1"));
        Assert.Equal(("Normal GL", "N", "N"), (row.GlType, row.VipInd, row.DepInd));
        Assert.Equal(1, await repo.UpdatePlanAsync(s, new("ZZ CIS WEB TEST", "ZZ1", "Normal GL", "T2.dot", "ZZ", "H", "Y", "D.dot", "r"), "Y"));
        Assert.Equal("T2.dot", Assert.Single(await repo.SearchAsync(s, "%ZZ CIS WEB TEST%", null, "ZZ1")).GlTemplate);
    }
}
