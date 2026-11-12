using Cis.Application.AcExport;
using Cis.Domain.Errors;
using Cis.Infrastructure.AcExport;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Payment Posting > Member rules (no database).</summary>
public class PostPayMbmRuleTests
{
    [Fact]
    public void A_pv_gives_7610_credit_and_82_group_debit()
    {
        var row = new Dictionary<string, object?> { ["PVNo"] = "CP100464273", ["GRPCoCode"] = "A014", ["MBMPaidAmt"] = 260m, ["PVPayto"] = "HATIBUL-UMAM BIN BASRI" };
        var r = PostPayMbmService.Lines([row], "G");
        Assert.Equal([("7610", "C", true), ("82A014", "D", false)], r.Lines.Select(l => (l.Account, l.DebitCredit, l.PostLine)).ToList());
        Assert.Equal("CP100464273", r.Lines[0].Reference);
        var d = PostPayMbmService.Lines([row], "D");
        Assert.Equal((260m, 0m), (d.Debit, d.Credit));
    }
}

/// <summary>Payment Posting > Member against the real databases (skipped without CIS_TEST_DB); the post is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PostPayMbmDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_post()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var service = new PostPayMbmService(TestDb.Sessions(), new PostPayMbmRepository(), new PostInvClnGstRepository(),
            new PostInvClnService(TestDb.Sessions(), new PostInvClnRepository(), exporter));
        var r = await service.SearchAsync(new PostPayMbmRequest { PvDateFrom = "01/07/2021", PvDateTo = "01/07/2021", PvFrom = "CP100464273", PvTo = "CP100464273" });
        Assert.Contains(r.Lines, l => l.Account == "82A014" && l.Amount == 260m);
        // A PV the desktop posted without its first character (batch 2654) is listed as posted.
        var legacy = await service.SearchAsync(new PostPayMbmRequest { Posted = "1 - Yes", PvDateFrom = "22/06/2021", PvDateTo = "22/06/2021", PvFrom = "CP100462686", PvTo = "CP100462686" });
        Assert.True(legacy.Posted);
        Assert.Equal(PostPayClnService.NoPeriod, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PostAsync(
            new PostPayMbmPostRequest(new PostPayMbmRequest { PvDateFrom = "01/07/2021", PvDateTo = "01/07/2021" }, " ", ["CP100464273"]), "DEVADMIN"))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        var batch = await new PostInvClnGstRepository().NextBatchAsync(tx);
        await new PostPayMbmRepository().InsertAsync(tx, "CP100464273", "07", "2021", batch, "DEVADMIN");
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISWorksheet.dbo.SUNPosting WHERE BatchNo = @batch AND TransRef = 'CP100464273' AND PostType = 'Pay2MBM'", new { batch }, tx.Sql().Transaction));
    }
}
