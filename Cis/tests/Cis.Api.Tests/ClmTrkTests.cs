using Cis.Application.Reports;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Reports;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Claim tracking catalogue and filter rules (no database).</summary>
public class ClmTrkRuleTests
{
    [Fact]
    public void The_catalogue_has_the_report_types_and_headers()
    {
        Assert.Equal(104, ClmTrkService.Catalogue.Count);
        Assert.All(ClmTrkService.Catalogue, e => Assert.NotEmpty(e.Headers));
        Assert.Contains(ClmTrkService.Catalogue, e => e.Type == 23);
    }

    [Fact]
    public void Only_known_filters_are_passed_typed()
    {
        var f = ClmTrkService.Filters(new() { ["payor"] = "DC", ["dovFrom"] = "01/07/2021", ["approvedFrom"] = "10", ["nope"] = "x" });
        Assert.Equal(["approvedFrom", "dovFrom", "payor"], f.Keys.Order(StringComparer.Ordinal));
        Assert.Equal(new DateTime(2021, 7, 1), f["dovFrom"]);
    }
}

/// <summary>Claim tracking against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class ClmTrkDatabaseTests
{
    [Fact]
    public async Task Ageband_summary_matches_the_desktop_procedure()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new ClmTrkService(TestDb.Sessions(), new ClmTrkRepository());
        var r = await service.SearchAsync(new ClmTrkRequest(1, "Summary", new() { ["payor"] = "DC", ["dovFrom"] = "01/07/2021", ["dovTo"] = "01/07/2021" }));
        await using var s = await TestDb.Sessions().OpenAsync();
        var legacy = (await s.Sql().Connection.QueryAsync("CISWorksheet.dbo.CISClaimTracking",
            new { strAppend = " AND PAYCode = 'DC' And DateVisit >= '07/01/2021' And DateVisit <= '07/01/2021'", strAppendCase = 14 },
            commandType: System.Data.CommandType.StoredProcedure, commandTimeout: 600)).ToList();
        Assert.Equal(legacy.Count, r.Rows.Count);
        Assert.Equal(legacy.Sum(x => (int)x.TotalCases), r.Rows.Sum(x => int.Parse(x[1])));
        // Diagnosis general summary (case chosen by the report choice) with the treatment reason descriptions.
        var diag = await service.SearchAsync(new ClmTrkRequest(7, "General Summary", new() { ["payor"] = "DC", ["dovFrom"] = "01/07/2021", ["dovTo"] = "01/07/2021" }));
        var legacyDiag = (await s.Sql().Connection.QueryAsync("CISWorksheet.dbo.CISClaimTracking",
            new { strAppend = " AND PAYCode = 'DC' And DateVisit >= '07/01/2021' And DateVisit <= '07/01/2021'", strAppendCase = 106 },
            commandType: System.Data.CommandType.StoredProcedure, commandTimeout: 600)).ToList();
        Assert.Equal(legacyDiag.Count, diag.Rows.Count);
        Assert.Contains(diag.Rows, x => x[1].Length > 0);
        // General listing without remarks (hand-ported row rules): one row per desktop case 93 row.
        var list = await service.SearchAsync(new ClmTrkRequest(9, "Without Remarks", new() { ["payor"] = "DC", ["dovFrom"] = "01/07/2021", ["dovTo"] = "01/07/2021" }));
        var legacyList = (await s.Sql().Connection.QueryAsync("CISWorksheet.dbo.CISClaimTracking",
            new { strAppend = " AND PAYCode = 'DC' And DateVisit >= '07/01/2021' And DateVisit <= '07/01/2021'", strAppendCase = 93 },
            commandType: System.Data.CommandType.StoredProcedure, commandTimeout: 600)).ToList();
        Assert.Equal(legacyList.Count, list.Rows.Count);
        Assert.Equal(54, list.Rows[0].Count);
        // O/S utilized (combined): consecutive rows of a member make one row, as the desktop's case 5 loop.
        var os = await service.SearchAsync(new ClmTrkRequest(16, "O/S Utilized (Combined)", new() { ["payor"] = "DC", ["dovFrom"] = "01/07/2021", ["dovTo"] = "01/07/2021" }));
        var legacyOs = (await s.Sql().Connection.QueryAsync("CISWorksheet.dbo.CISClaimTracking",
            new { strAppend = " AND PAYCode = 'DC' And DateVisit >= '07/01/2021' And DateVisit <= '07/01/2021'", strAppendCase = 5 },
            commandType: System.Data.CommandType.StoredProcedure, commandTimeout: 600)).Select(x => (string)x.MBMNumber).ToList();
        Assert.Equal(legacyOs.Where((m, i) => i == 0 || m != legacyOs[i - 1]).Count(), os.Rows.Count);
        // Admission reports: the desktop's checks, and the monthly report's payors (case 35) with their days.
        Assert.Equal(ClmTrkService.SameMonth, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new ClmTrkRequest(0, "Monthly Report",
            new() { ["dovFrom"] = "30/06/2021", ["dovTo"] = "01/07/2021" })))).Message);
        Assert.Equal(ClmTrkService.NoDov, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new ClmTrkRequest(0, "Monthly Report", new())))).Message);
        Assert.Equal(ClmTrkService.NoYear, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new ClmTrkRequest(0, "YTD Report", new())))).Message);
        var monthly = await service.SearchAsync(new ClmTrkRequest(0, "Monthly Report", new() { ["payor"] = "DC", ["dovFrom"] = "01/07/2021", ["dovTo"] = "03/07/2021" }));
        Assert.Equal((await s.Sql().Connection.QueryAsync("CISWorksheet.dbo.CISClaimTracking",
            new { strAppend = " AND PAYCode = 'DC' And DateVisit >= '07/01/2021' And DateVisit <= '07/03/2021'", strAppendCase = 35 },
            commandType: System.Data.CommandType.StoredProcedure, commandTimeout: 600)).Count(), monthly.Rows.Count);
        Assert.Matches(@"^\d{2}/07/2021-\w+day$", monthly.Rows[0][1]);
        Assert.Equal(MbmTrkService.Pending, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SearchAsync(new ClmTrkRequest(4, "By Clinic Type & Clinic Panel", new())))).Message);
    }
}
