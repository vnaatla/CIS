using Cis.Application.Reports;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Reports;
using Dapper;

namespace Cis.Api.Tests;

/// <summary>Member tracking catalogue and filter rules (no database).</summary>
public class MbmTrkRuleTests
{
    [Fact]
    public void The_catalogue_has_every_report_type_and_its_headers()
    {
        Assert.Equal(73, MbmTrkService.Catalogue.Count);
        Assert.All(MbmTrkService.Catalogue, e => Assert.NotEmpty(e.Headers));
        Assert.Equal(Enumerable.Range(0, 18), MbmTrkService.Catalogue.Select(e => e.Type).Distinct().Order());
    }

    [Fact]
    public void Only_known_filters_are_passed_typed()
    {
        var f = MbmTrkService.Filters(new() { ["payor"] = "DC", ["effFrom"] = "01/07/2021", ["hltBlank"] = "true", ["evil; drop"] = "x", ["premiumFrom"] = "10.5" });
        Assert.Equal(["effFrom", "hltBlank", "payor", "premiumFrom"], f.Keys.Order(StringComparer.Ordinal));
        Assert.Equal(new DateTime(2021, 7, 1), f["effFrom"]);
        Assert.Equal(true, f["hltBlank"]);
        Assert.Equal(10.5m, f["premiumFrom"]);
    }
}

/// <summary>Member tracking against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class MbmTrkDatabaseTests
{
    [Fact]
    public async Task Ins_type_by_payor_matches_the_desktop_procedure()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new MbmTrkService(TestDb.Sessions(), new MbmTrkRepository());
        var r = await service.SearchAsync(new MbmTrkRequest(10, "By Payor", "P", new() { ["payor"] = "DC", ["status"] = "0" }));
        await using var s = await TestDb.Sessions().OpenAsync();
        var legacy = (await s.Sql().Connection.QueryAsync("HISDB.dbo.CISMBMTracking", new { strAppend = " AND PAYCode = 'DC' And MBMStatus = 'A' ", strAppendCase = 113 },
            commandType: System.Data.CommandType.StoredProcedure, commandTimeout: 600)).ToList();
        Assert.Equal(legacy.Count, r.Rows.Count);
        Assert.Equal(legacy.Sum(x => (int)x.TotalP), r.Principals);
        // A month report (Eff. Month by payor, principals): the same rows as the desktop's case 117, labelled year / zero-padded month.
        var month = await service.SearchAsync(new MbmTrkRequest(11, "Eff. Month (By Payor)", "P", new() { ["payor"] = "DC", ["status"] = "0" }));
        var legacyMonth = (await s.Sql().Connection.QueryAsync("HISDB.dbo.CISMBMTracking", new { strAppend = " AND PAYCode = 'DC' And MBMStatus = 'A' ", strAppendCase = 117 },
            commandType: System.Data.CommandType.StoredProcedure, commandTimeout: 600)).ToList();
        Assert.Equal(legacyMonth.Count, month.Rows.Count);
        Assert.Matches(@"^\d{4}/\d{2}$", month.Rows[0][1]);
        Assert.Equal(MbmTrkService.NoReport, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SearchAsync(new MbmTrkRequest(10, "Nope", "P", new()))))
            .Message);
    }
}
