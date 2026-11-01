using Cis.Application.Maintenance.Plan;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Maintenance;
using Dapper;

namespace Cis.Api.Tests;

public class PlanAnnualLimitServiceTests
{
    internal static readonly AnnualLimitDto Valid = new(null, "A001 - PLAN A", "N - NON-SIHAT", "H - GROUP FAMILY", null, null, "1,000.00", "200", "150",
        "50", "A", "", "", "", "", "01/01/2026", "3");

    private readonly PlanAnnualLimitService _service = new(new FakeSessions(), null!, new FixedClock(new DateTime(2026, 11, 1)));

    [Fact]
    public async Task Save_checks_in_the_desktop_order()
    {
        Assert.Equal("Please Enter Plan Code", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { PlnCode = "" }, false, "DEV"))).Message);
        Assert.Equal("Please Enter Principal MEP Annual Limit", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { PMep = " " }, false, "DEV"))).Message);
        Assert.Equal("Please Enter Annual Limit Effective Date", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { EffDate = "__/__/____" }, false, "DEV"))).Message);
        Assert.Equal("Please Enter Supp Premium Limit", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { SuppStatus = "B" }, false, "DEV"))).Message);
        Assert.Equal("Please Enter Insured Code", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.DeleteAsync(Valid with { PlnCode = "" }))).Message);
    }

    [Fact]
    public void Values_follow_SaveANN()
    {
        var r = PlanAnnualLimitService.Row(Valid, false);
        Assert.Equal(("A001", "H", "N", 1000m, "A", (decimal?)null, 3), (r.PlnCode, r.InsCode, r.HltCode, r.AnnLimit, r.Status, r.SuppLimit, r.Version));
        var c = PlanAnnualLimitService.Row(Valid with { SuppStatus = "C - Separate", SuppLimit = "10", SDental = "1", SSpec = "2", SMep = "3", Index = 7 }, true);
        Assert.Equal((7, "C", 10m, 3m), (c.Index, c.Status, c.SuppLimit, c.SMep));
    }
}

/// <summary>Annual Limit procedures against the real CISDB (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PlanAnnualLimitDatabaseTests
{
    [Fact]
    public async Task Annual_limit_is_saved_refused_twice_updated_and_deleted()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new PlanAnnualLimitRepository();
        Assert.NotEmpty((await repo.SetupAsync(s)).Plans);
        var row = PlanAnnualLimitService.Row(PlanAnnualLimitServiceTests.Valid with { PlnCode = "ZZ9", EffDate = "01/01/1901" }, false);
        Assert.False(await repo.ExistsAsync(s, row));
        Assert.Equal(1, await repo.SaveAsync(s, row, "CISTEST", new DateTime(2026, 11, 1)));
        Assert.True(await repo.ExistsAsync(s, row));
        var index = await s.Sql().Connection.ExecuteScalarAsync<int>("SELECT MAX(AnnualIndex) FROM dbo.AnnualLimit WHERE PLNCode = 'ZZ9'", transaction: s.Sql().Transaction);
        Assert.Equal(1, await repo.SaveAsync(s, row with { Index = index, AnnLimit = 2000m }, "CISTEST", new DateTime(2026, 11, 1)));
        Assert.False(await repo.InUseAsync(s, "ZZ9"));
        Assert.Equal(1, await repo.DeleteAsync(s, index));
    }
}
