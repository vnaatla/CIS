using Cis.Application.Reports;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Reports;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Claim Tracking 2 catalogue and filter rules (no database).</summary>
public class ClmTrk2RuleTests
{
    [Fact]
    public void The_catalogue_has_both_amount_bases_and_cost_detail()
    {
        var service = new ClmTrk2Service(null!, null!);
        var reports = service.Reports();
        Assert.Equal(35, reports.Count);
        Assert.Equal(17, reports.Count(r => r.Basis == 1));
        Assert.Contains(reports, r => r.Basis == 3 && r.Name == "Cost Detail");
        Assert.All(ClmTrk2Service.Catalogues["main"], e => Assert.Equal(e.Cols.Count, e.Headers.Count));
    }

    [Fact]
    public void The_jupiter_copy_compares_texts_exactly()
    {
        Assert.False(ClmTrk2Service.Filters(new() { ["agent"] = "A1", ["exact"] = "true" }, "main").ContainsKey("exact"));
        Assert.Equal(true, ClmTrk2Service.Filters(new() { ["agent"] = "A1" }, "jupiter")["exact"]);
    }
}

/// <summary>Claim Tracking 2 against the real databases (skipped without CIS_TEST_DB); read-only. The desktop procedure itself fails (double-quoted values under
/// QUOTED_IDENTIFIER ON), so the figures are checked against the views it reads.</summary>
[Collection(DatabaseCollection.Name)]
public class ClmTrk2DatabaseTests
{
    [Fact]
    public async Task Actual_claim_listing_and_cost_detail_match_the_views()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new ClmTrk2Service(TestDb.Sessions(), new ClmTrk2Repository());
        var day = new Dictionary<string, string?> { ["payor"] = "EA", ["dovFrom"] = "01/07/2021", ["dovTo"] = "01/07/2021" };
        await using var s = await TestDb.Sessions().OpenAsync();
        const string where = " WHERE SUBSTRING(PAYCode,1,1) <> 'Z' AND PAYCode = 'EA' AND DateVisit >= '20210701' AND DateVisit <= '20210701'";
        var claims = await service.SearchAsync(new ClmTrk2Request(1, 4, day));
        Assert.Equal(await s.Sql().Connection.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM CISWorksheet.dbo.VIEW_CLMReport_Actual" + where), claims.Rows.Count);
        Assert.Equal(32, claims.Rows[0].Count);
        var cost = await service.SearchAsync(new ClmTrk2Request(3, 0, day));
        Assert.Equal(29, cost.Rows.Count);
        var actual = await s.Sql().Connection.ExecuteScalarAsync<decimal?>("SELECT SUM(" + string.Join(" + ", ClmTrk2Service.Fees.Select(f => $"ISNULL({f}, 0)")) + ") FROM CISWorksheet.dbo.VIEW_CLMReport_Actual" + where) ?? 0;
        Assert.Equal(actual.ToString("#,##0.00", System.Globalization.CultureInfo.InvariantCulture), cost.Rows[28][1]);
    }
}
