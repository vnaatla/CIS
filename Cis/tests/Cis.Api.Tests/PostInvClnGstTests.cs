using Cis.Application.AcExport;
using Cis.Domain.Errors;
using Cis.Infrastructure.AcExport;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Invoice Posting > Clinic (GST) rules (no database).</summary>
public class PostInvClnGstRuleTests
{
    [Fact]
    public void An_ffs_invoice_gives_four_lines_with_the_gst_split()
    {
        var row = new Dictionary<string, object?>
        {
            ["InvoiceNo"] = "059809", ["ClinicCode"] = "D153", ["CLMPayableByMedix2C"] = 608.855m, ["CLMTotalActual"] = 640.90m, ["CLMCoPayColl"] = 0m,
            ["CLMMCOProcFee"] = 32.045m,
        };
        var r = PostInvClnGstService.Lines([row], "G", new DateTime(2026, 11, 9));
        Assert.Equal([("6000", "D", 640.90m), ("84D153", "C", 608.86m), ("1002", "C", 30.23m), ("8830", "C", 1.82m)],
            r.Lines.Select(l => (l.Account, l.DebitCredit, l.Amount)).ToList());
        Assert.True(r.Lines[0].PostLine);
        Assert.Equal((608.86m, 3 * 608.86m), (r.Debit, r.Credit));
        Assert.Equal((0m, 0m), PostInvClnGstService.Split(32.045m, new DateTime(2015, 3, 31)));
        Assert.Equal(3, PostInvClnGstService.Lines([row], "G", new DateTime(2015, 3, 31)).Lines.Count);
    }
}

/// <summary>Invoice Posting > Clinic (GST) against the real databases (skipped without CIS_TEST_DB); the post is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PostInvClnGstDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_post()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var clinic = new PostInvClnService(TestDb.Sessions(), new PostInvClnRepository(), exporter);
        var service = new PostInvClnGstService(TestDb.Sessions(), new PostInvClnGstRepository(), new PostInvClnRepository(), clinic, TimeProvider.System);
        var search = new PostInvClnRequest { InvDateFrom = "04/08/2017", InvDateTo = "04/08/2017", InvoiceFrom = "059809", InvoiceTo = "059809" };
        var r = await service.SearchAsync(search);
        Assert.Equal(["6000", "84D153", "1002", "8830"], r.Lines.Select(l => l.Account).ToArray());
        Assert.Equal(PostInvClnGstService.NoInvDate, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PostAsync(
            new PostInvClnPostRequest(new PostInvClnRequest(), "08/2017", [new("059809", "D153", new DateTime(2017, 8, 4))]), "DEVADMIN"))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        var repo = new PostInvClnGstRepository();
        var batch = await repo.NextBatchAsync(tx);
        Assert.Matches("^[0-9]+$", batch);
        await repo.InsertAsync(tx, new PostInvClnTick("059809", "D153", new DateTime(2017, 8, 4)), "08", "2017", "TTEST01", batch, "DEVADMIN");
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISWorksheet.dbo.SUNPosting WHERE BatchNo = @batch AND CLNInvoiceBatch = 'TTEST01'", new { batch }, tx.Sql().Transaction));
    }
}
