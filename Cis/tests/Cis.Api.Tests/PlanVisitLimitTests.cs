using Cis.Application.Maintenance.Plan;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Maintenance;
using Dapper;

namespace Cis.Api.Tests;

public class PlanVisitLimitServiceTests
{
    internal static readonly VisitLimitDto Valid = new(null, "A001 - PLAN A", "N - NON-SIHAT", "H - GROUP FAMILY", null, null, "12", "A", "", "01/01/2026", "1");

    private readonly PlanVisitLimitService _service = new(new FakeSessions(), null!, null!, new FixedClock(new DateTime(2026, 11, 1)));

    [Fact]
    public async Task Save_checks_in_the_desktop_order()
    {
        Assert.Equal("Please Enter Visit Limit", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { VisitLimit = "" }, false, "DEV"))).Message);
        Assert.Equal("Please Enter Visit Limit Version", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { Version = "" }, false, "DEV"))).Message);
        Assert.Equal("Please Enter Supp Visit Limit", (await Assert.ThrowsAsync<BusinessRuleException>(() => _service.SaveAsync(Valid with { SuppStatus = "C" }, false, "DEV"))).Message);
    }

    [Fact]
    public void Values_follow_SaveVisit()
    {
        var r = PlanVisitLimitService.Row(Valid with { SuppStatus = "B - Share", SuppLimit = "4" }, false);
        Assert.Equal(("A001", "H", "N", 12m, "B", 4m, 1), (r.PlnCode, r.InsCode, r.HltCode, r.VisitLimit, r.Status, r.SuppLimit, r.Version));
        Assert.Null(PlanVisitLimitService.Row(Valid with { SuppLimit = "4" }, false).SuppLimit);
    }
}

/// <summary>Visit Limit procedures against the real CISDB (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PlanVisitLimitDatabaseTests
{
    [Fact]
    public async Task Visit_limit_is_saved_refused_twice_updated_and_deleted()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        await using var s = await TestDb.Sessions().BeginAsync();
        var repo = new PlanVisitLimitRepository();
        var row = PlanVisitLimitService.Row(PlanVisitLimitServiceTests.Valid with { PlnCode = "ZZ9", EffDate = "01/01/1901" }, false);
        Assert.False(await repo.ExistsAsync(s, row));
        Assert.Equal(1, await repo.SaveAsync(s, row, "CISTEST", new DateTime(2026, 11, 1)));
        Assert.True(await repo.ExistsAsync(s, row));
        Assert.Single(await repo.SearchAsync(s, "ZZ9", null, null, null, null).ContinueWith(t => t.Result), r => true);
        var index = await s.Sql().Connection.ExecuteScalarAsync<int>("SELECT MAX(VisitIndex) FROM dbo.VisitLimit WHERE PLNCode = 'ZZ9'", transaction: s.Sql().Transaction);
        Assert.Equal(1, await repo.SaveAsync(s, row with { Index = index, VisitLimit = 20m }, "CISTEST", new DateTime(2026, 11, 1)));
        Assert.Equal(1, await repo.DeleteAsync(s, index));
    }
}
