using Cis.Application.Maintenance.Plan;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Maintenance;

namespace Cis.Api.Tests;

public class PlanCodeServiceTests
{
    internal static readonly PlanCodeDto Valid = new("N - NON-SIHAT", "DC - DIRECT CLIENT", "AA - SOMETHING", "ZZ99 - ZZ TEST CO", "CSMD A", "Yes", "No",
        "Yes", "No", "Yes", "No", "Y - Yes", "2", ["", ""], ["first plan", "second plan"], "Y-Yes", "N-No", "B-Both", "100", "x", "", "50.5",
        "note", "N - No");

    private readonly PlanCodeService _service = new(new FakeSessions(), null!, new FixedClock(new DateTime(2026, 11, 1)));

    [Fact]
    public async Task Save_checks_run_in_the_desktop_order()
    {
        Assert.Equal("Please Enter Health Code ", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { HealthType = "" }, "DEV"))).Message);
        Assert.Equal("Please Enter Payor Code ", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { PayCode = null }, "DEV"))).Message);
        Assert.Equal("Please Enter Supplementary Visit Limit", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { SVisLmt = "" }, "DEV"))).Message);
        Assert.Equal("Please Enter Plan Code", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { HealthType = "S - SIHAT" }, "DEV"))).Message);
        Assert.Equal(PlanCodeService.OnlyNonSihat, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { HealthType = "S - SIHAT", PlanCodes = ["S01"] }, "DEV"))).Message);
        Assert.Equal("Please Select OP Dependent", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { OpDependent = " " }, "DEV"))).Message);
    }

    [Fact]
    public void Values_and_filters_follow_the_desktop()
    {
        var r = PlanCodeService.Row(Valid, "Z001", "first plan");
        Assert.Equal(("FIRST PLAN", "N", "DC", "AA", "ZZ99", "Y", "N", "Y", "Y", "N", "B-Both"),
            (r.Desc, r.HltCode, r.PayCode, r.HisPayCode, r.GrpCompany, r.AnnLmt, r.VisLmt, r.DcStatus, r.Insured, r.HrApproval, r.OpFamilyStatus));
        Assert.Equal((100d, 0d, 0d, 50.5d, "NOTE", "N"), (r.PrinSpLimit, r.PrinGpLimit, r.DepSpLimit, r.DepGpLimit, r.SpecialRemark, r.OpDependent));
        var f = PlanCodeService.Reduce(Valid with { PlanCodes = ["a0"] });
        Assert.Equal(("A0", "DC", "AA", "N", "ZZ99", "Y", "N"), (f.PlnCode, f.PayCode, f.HisPayCode, f.HltCode, f.GrpCode, f.Flags["annLmt"], f.Flags["visLmt"]));
        Assert.Null(f.Flags["pDentalStatus"]);
    }
}

/// <summary>Plan Code procedures against the real CISDB (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PlanDatabaseTests
{
    [Fact]
    public async Task Plan_code_is_generated_saved_found_by_its_indicators_and_deleted()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new PlanCodeRepository();
        var setup = await repo.SetupAsync(s);
        Assert.NotEmpty(setup.HealthTypes);
        Assert.NotEmpty(setup.ClinicContracts);
        Assert.NotEmpty(await repo.PayorsAsync(s, "N"));
        var group = setup.GroupCompanies.First(g => g.Code.Length == 4).Code;
        var code = await repo.NextCodeAsync(s, group[..1]);
        Assert.Matches("^[A-Z0-9][0-9]{3}$", code);
        var plan = PlanCodeServiceTests.Valid with { GroupCompany = group + " - X", ClnContract = setup.ClinicContracts[0].Code };
        await repo.InsertAsync(s, PlanCodeService.Row(plan, code, "ZZ CIS WEB PLAN"), "CISTEST", new DateTime(2026, 11, 1));
        var found = Assert.Single(await repo.SearchAsync(s, PlanCodeService.Reduce(plan with { PlanCodes = [code] })));
        Assert.Equal("ZZ CIS WEB PLAN", found["PLNDesc"]);
        Assert.Empty(await repo.SearchAsync(s, PlanCodeService.Reduce(plan with { PlanCodes = [code], AnnLmt = "No" })));
        Assert.Equal("Y", (await repo.GetAsync(s, code))!["AnnLmt"]);
        Assert.False(await repo.InUseAsync(s, code));
        Assert.Equal(1, await repo.DeleteAsync(s, code));
    }
}
