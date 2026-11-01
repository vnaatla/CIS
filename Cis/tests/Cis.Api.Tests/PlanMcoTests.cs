using Cis.Application.Maintenance.Plan;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Maintenance;

namespace Cis.Api.Tests;

public class PlanMcoServiceTests
{
    internal static readonly McoDto Valid = new("N - NON-SIHAT", "A001", "A", "01/01/2026",
        new() { ["IA"] = "10", ["IC"] = "", ["FA"] = "x", ["SFA"] = "3", ["HA"] = "0" });

    private readonly PlanMcoService _service = new(new FakeSessions(), null!, new FixedClock(new DateTime(2026, 11, 1)));

    [Fact]
    public async Task Save_checks_in_the_desktop_order()
    {
        Assert.Equal(PlanMcoService.HltFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { HltCode = "" }, "DEV"))).Message);
        Assert.Equal(PlanMcoService.PlnFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { PlnCode = " " }, "DEV"))).Message);
        Assert.Equal(PlanMcoService.FeesFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { Fees = new() { ["SFA"] = "3" } }, "DEV"))).Message);
        Assert.Equal(PlanMcoService.BadDate, (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { EffDate = "__/__/____" }, "DEV"))).Message);
    }

    [Fact]
    public void Only_types_with_an_amount_are_written()
    {
        var rows = PlanMcoService.Rows(Valid);
        Assert.Equal(["I", "F"], rows.Select(r => r.InsCode));
        Assert.Equal((10d, 0d, 0d, 3d), (rows[0].Adult, rows[0].Child, rows[1].Adult, rows[1].SuppAdult));
    }
}

/// <summary>MCO procedures against the real CISDB (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PlanMcoDatabaseTests
{
    [Fact]
    public async Task Mco_fees_are_saved_found_refused_twice_and_deleted()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new PlanMcoRepository();
        Assert.NotEmpty((await repo.SetupAsync(s)).Plans);
        var eff = new DateTime(1901, 1, 1);
        Assert.False(await repo.ExistsAsync(s, "ZZ9", "N", eff, "I"));
        await repo.InsertAsync(s, "ZZ9", "N", "A", eff, new McoRow("I", 10, 5, 0, 0), "CISTEST", new DateTime(2026, 11, 1));
        Assert.True(await repo.ExistsAsync(s, "ZZ9", "N", eff, "I"));
        var row = Assert.Single(await repo.SearchAsync(s, "N", "ZZ9", eff));
        Assert.Equal(1, await repo.DeleteAsync(s, Convert.ToInt32(row["MCOIndex"])));
    }
}
