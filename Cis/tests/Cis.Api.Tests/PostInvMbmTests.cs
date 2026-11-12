using Cis.Application.AcExport;
using Cis.Domain.Errors;
using Cis.Infrastructure.AcExport;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Invoice Posting > Member rules (no database).</summary>
public class PostInvMbmRuleTests
{
    [Fact]
    public void A_case_with_postage_gives_debit_group_and_postage_lines()
    {
        var row = new Dictionary<string, object?>
        {
            ["CASNumber"] = "EA1", ["MBMNumber"] = "UBHA0332351*06", ["GRPCoCode"] = "U063", ["CLMTotalActual"] = 178m, ["CLMPostageFee"] = 3m, ["PatientName"] = "NUR",
        };
        var g = PostInvMbmService.Lines([row], "G");
        Assert.Equal([("6000", "D", 178m, false), ("82U063", "C", 175m, false), ("4040", "C", 3m, true)],
            g.Lines.Select(l => (l.Account, l.DebitCredit, l.Amount, l.PostLine)).ToList());
        Assert.Equal("2351", g.Lines[0].Extra!["Member"]);
        Assert.Equal([("82U063", 175m), ("4040", 3m)], PostInvMbmService.Lines([row], "C").Lines.Select(l => (l.Account, l.Amount)).ToList());
        Assert.Single(PostInvMbmService.Lines([row], "D").Lines);
    }
}

/// <summary>Invoice Posting > Member against the real databases (skipped without CIS_TEST_DB); the post is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PostInvMbmDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_post()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var service = new PostInvMbmService(TestDb.Sessions(), new PostInvMbmRepository(), new PostInvClnGstRepository(),
            new PostInvClnService(TestDb.Sessions(), new PostInvClnRepository(), exporter));
        Assert.NotEmpty(await service.CompaniesAsync());
        var r = await service.SearchAsync(new PostInvMbmRequest { InvDateFrom = "01/07/2021", InvDateTo = "01/07/2021", CaseFrom = "EA12538622", CaseTo = "EA12538622" });
        Assert.Contains(r.Lines, l => l.Account == "82U063" && l.Description.StartsWith("Nur Amilin", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(PostInvMbmService.NoPeriod, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PostAsync(new PostInvMbmPostRequest(
            new PostInvMbmRequest { BordxDateFrom = "01/07/2021", BordxDateTo = "31/07/2021" }, null, [new("EA12538622", null)]), "DEVADMIN"))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        var batch = await new PostInvClnGstRepository().NextBatchAsync(tx);
        await new PostInvMbmRepository().InsertAsync(tx, new PostInvMbmTick("EA12538622", null), "07", "2021", batch, "DEVADMIN");
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISWorksheet.dbo.SUNPosting WHERE BatchNo = @batch AND TransRef = 'EA12538622' AND PostType = 'InvMBM'", new { batch }, tx.Sql().Transaction));
    }
}
