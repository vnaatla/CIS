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
    public void The_jupiter_screen_has_its_own_choices_and_columns()
    {
        var jupiter = MbmTrkService.Catalogues["jupiter"];
        Assert.Equal(72, jupiter.Count);
        Assert.DoesNotContain(jupiter, e => e.Selection == "Card Sent Date Listing");
        // Its detail listings show the supplementaries' payor effective date (the main form: their cover effective date).
        Assert.Contains(jupiter.First(e => e.Routine == "GeneralListing4").Detail!.Supp, c => c.N == "MBMPayorEffDate");
        Assert.Contains(MbmTrkService.Catalogue.First(e => e.Routine == "GeneralListing4").Detail!.Supp, c => c.N == "MBMCEffDate");
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
        // A detail listing (General Listing, principals of one member): the same rows as the desktop's case 88.
        var detail = await service.SearchAsync(new MbmTrkRequest(7, "General Listing", "P", new() { ["mbmNo"] = "WKHA3578533*01" }));
        var legacyDetail = (await s.Sql().Connection.QueryAsync("HISDB.dbo.CISMBMTracking", new { strAppend = " AND MBMNumber = 'WKHA3578533*01'", strAppendCase = 88 },
            commandType: System.Data.CommandType.StoredProcedure, commandTimeout: 600)).ToList();
        Assert.Equal(legacyDetail.Select(x => (string)x.MBMNumber).Distinct().Count(), detail.Rows.Count);
        Assert.Equal("WKHA3578533*01", detail.Rows[0][0]);
        // Summaries ported by hand. Bordx MCO (INS & MBM type): one row per desktop case 37 group, the H column summed from case 44.
        var dc = new Dictionary<string, string?> { ["payor"] = "DC", ["status"] = "0" };
        var bordx = await service.SearchAsync(new MbmTrkRequest(2, "INS Type & MBM Type (MCO)", "P", dc));
        async Task<List<dynamic>> Legacy(int c, string append) => (await s.Sql().Connection.QueryAsync("HISDB.dbo.CISMBMTracking", new { strAppend = append, strAppendCase = c },
            commandType: System.Data.CommandType.StoredProcedure, commandTimeout: 600)).ToList();
        const string dcActive = " AND PAYCode = 'DC' And MBMStatus = 'A' ";
        Assert.Equal((await Legacy(37, dcActive)).Count, bordx.Rows.Count);
        Assert.Equal((await Legacy(44, dcActive)).Sum(x => (int)x.TotalH), bordx.Rows.Sum(r => int.Parse(r[16])));
        // MCO invoice summary (Both): the desktop's case 178 rows.
        var inv = await service.SearchAsync(new MbmTrkRequest(12, "Summary", "B", new() { ["payor"] = "DC", ["invoiceDateFrom"] = "01/07/2021", ["invoiceDateTo"] = "31/07/2021" }));
        Assert.Equal((await Legacy(178, " AND PAYCode = 'DC' AND MBMInvoiceDate >= '07/01/2021' AND MBMInvoiceDate <= '07/31/2021' ")).Count, inv.Rows.Count);
        Assert.All(service.Reports(), x => Assert.True(x.Migrated, $"{x.Type} {x.Selection}"));
        // Branch by agent (payor DC, registered 01/07/2021): the desktop's case 61 groups; individual policies summed from case 66 with INSCode I.
        var oneDay = new Dictionary<string, string?> { ["payor"] = "DC", ["dorFrom"] = "01/07/2021", ["dorTo"] = "01/07/2021" };
        const string oneDayLegacy = " AND PAYCode = 'DC' AND MBMValueDate >= '07/01/2021' AND MBMValueDate <= '07/01/2021' ";
        var branch = await service.SearchAsync(new MbmTrkRequest(3, "By Agent", "P", oneDay));
        Assert.Equal((await Legacy(61, oneDayLegacy)).Count, branch.Rows.Count);
        Assert.Equal((await Legacy(66, oneDayLegacy + " AND INSCode = 'I'")).Sum(x => (int)x.TotalP), branch.Rows.Sum(r => int.Parse(r[2])));
        Assert.Equal(MbmTrkService.NoReport, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SearchAsync(new MbmTrkRequest(10, "Nope", "P", new()))))
            .Message);
    }
}
