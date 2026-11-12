using Cis.Application.Reports;
using Cis.Domain.Errors;
using Cis.Infrastructure.Reports;

namespace Cis.Api.Tests;

/// <summary>Invoice tracking column rules (no database).</summary>
public class InvTrkRuleTests
{
    [Fact]
    public void Derived_columns_follow_the_desktop()
    {
        var x = new Dictionary<string, object?>
        {
            ["NoofDays"] = "0", ["INVPayTo"] = "M", ["CLNName"] = "PANEL", ["NPClinic"] = "OTHER", ["CASTreatReason5"] = "199", ["CASDiagnosis5"] = "FEVER",
            ["ReasonDesc5"] = "R5", ["GPFee"] = 10m, ["SPFee"] = null, ["GPMed"] = 2m, ["SPInj"] = 1.5m,
        };
        InvTrkService.Shape(x);
        Assert.Equal(("1", "OTHER", "FEVER", 10m, 2m, 1.5m), (x["DaysText"], x["ClinicText"], x["Reason5Text"], x["FeeTotal"], x["MedTotal"], x["InjTotal"]));
    }
}

/// <summary>Invoice tracking against the real databases (skipped without CIS_TEST_DB); read-only.</summary>
[Collection(DatabaseCollection.Name)]
public class InvTrkDatabaseTests
{
    [Fact]
    public async Task Lookups_search_and_checks()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new InvTrkService(TestDb.Sessions(), new InvTrkRepository());
        Assert.NotEmpty(await service.LookupAsync("companies", "DC - x"));
        var r = await service.SearchAsync(new InvTrkRequest { CaseFrom = "EA12538600", CaseTo = "EA12538699" });
        Assert.True(r.Rows.Count > 50);
        Assert.All(r.Rows, x => Assert.InRange(string.CompareOrdinal(((string)x["CASNumber"]!).Trim(), "EA12538600"), 0, int.MaxValue));
        Assert.True(r.Total > 0);
        var j = await service.SearchAsync(new InvTrkRequest { CaseFrom = "EA12538600", CaseTo = "EA12538699", Status = "O - Open" }, true);
        Assert.All(j.Rows, x => Assert.True(x.ContainsKey("MSNO")));
        Assert.Equal(InvTrkService.BadDate, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new InvTrkRequest { VisitFrom = "31/02/2021" }))).Message);
        Assert.Equal(InvTrkService.NoRecord, (await Assert.ThrowsAsync<NotFoundException>(() => service.SearchAsync(new InvTrkRequest { CaseFrom = "ZZ", CaseTo = "ZZ" }))).Message);
    }
}
