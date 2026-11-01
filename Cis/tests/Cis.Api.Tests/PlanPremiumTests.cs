using Cis.Application.Maintenance.Plan;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Maintenance;

namespace Cis.Api.Tests;

public class PlanPremiumServiceTests
{
    internal static readonly PremiumDto Valid = new("H - GROUP FAMILY", "N - NON-SIHAT", "DC - DIRECT CLIENT", "1 - AGE", "A001 - PLAN A", "01/01/2026",
        "2", "A", ["100", "", "1,250.50", "", "", "", ""], ["9", "", "", "", "", "", ""]);

    private readonly PlanPremiumService _service = new(new FakeSessions(), null!, new FixedClock(new DateTime(2026, 11, 1)));

    [Theory]
    [InlineData("", "N", "A001", "01/01/2026", PlanPremiumService.InsFirst)]
    [InlineData("H", " ", "A001", "01/01/2026", PlanPremiumService.HltFirst)]
    [InlineData("H", "N", "", "01/01/2026", PlanPremiumService.PlnFirst)]
    [InlineData("H", "N", "A001", "__/__/____", PlanPremiumService.BadDate)]
    public async Task Save_checks_in_the_desktop_order(string ins, string hlt, string pln, string eff, string message) =>
        Assert.Equal(message, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            _service.SaveAsync(Valid with { InsCode = ins, HltCode = hlt, PlnCode = pln, EffDate = eff }, "DEV"))).Message);

    [Fact]
    public void Seven_rows_like_SavePRM()
    {
        var a = PlanPremiumService.Rows(Valid, new DateTime(2026, 1, 1));
        Assert.Equal(7, a.Count);
        Assert.Equal(("H", "DC", "N", "01", "A001", 100m, (decimal?)null, "A", 2), (a[0].InsCode, a[0].PayCode, a[0].HltCode, a[0].AgeCode, a[0].PlnCode, a[0].Amount, a[0].SuppAmount, a[0].SuppStatus, a[0].Version));
        Assert.Equal(("1", (decimal?)null), (a[1].AgeCode, a[1].Amount));
        Assert.Equal(("03", 1250.50m), (a[2].AgeCode, a[2].Amount));
        var b = PlanPremiumService.Rows(Valid with { SuppStatus = "B - Share Limit for Supplementary" }, new DateTime(2026, 1, 1));
        Assert.Equal((9m, 0m, (decimal?)null), (b[0].SuppAmount, b[2].SuppAmount, b[1].SuppAmount));
    }
}

/// <summary>Premium procedures against the real CISDB (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PlanPremiumDatabaseTests
{
    [Fact]
    public async Task Premium_rows_are_saved_found_refused_twice_and_deleted()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new PlanPremiumRepository();
        var setup = await repo.SetupAsync(s);
        Assert.NotEmpty(setup.Plans);
        var rows = PlanPremiumService.Rows(PlanPremiumServiceTests.Valid with { PlnCode = "ZZ9", EffDate = "01/01/1901" }, new DateTime(1901, 1, 1));
        Assert.False(await repo.ExistsAsync(s, rows[0], "DC"));
        foreach (var r in rows) await repo.InsertAsync(s, r, "CISTEST", new DateTime(2026, 11, 1));
        Assert.True(await repo.ExistsAsync(s, rows[0], "DC"));
        var found = await repo.SearchAsync(s, "H", null, null, null, "ZZ9");
        Assert.Equal(7, found.Count);
        Assert.Contains(await repo.BandsAsync(s, "H", "ZZ9"), b => b.AgeCode == "03" && b.Amount == 1250.50m);
        Assert.False(await repo.InUseAsync(s, "H", "N", "ZZ9", "01", "DC"));
        Assert.Equal(1, await repo.DeleteAsync(s, Convert.ToInt32(found[0]["PRMCode"])));
    }
}
