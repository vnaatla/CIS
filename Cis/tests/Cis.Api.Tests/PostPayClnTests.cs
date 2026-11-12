using Cis.Application.AcExport;
using Cis.Domain.Errors;
using Cis.Infrastructure.AcExport;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Payment Posting > Clinic rules (no database).</summary>
public class PostPayClnRuleTests
{
    [Fact]
    public void A_pv_gives_a_7610_credit_and_an_84_clinic_debit()
    {
        var row = new Dictionary<string, object?> { ["PVCLNNo"] = "CP1", ["ClinicCode"] = "B294", ["CLNName"] = "KLINIK BALA", ["CLNPaidAmt"] = 34.2m, ["PayCode"] = "DC" };
        var r = PostPayClnService.Lines([row], "G");
        Assert.Equal([("7610", "C", "UHC", false), ("84B294", "D", "UHC", true)], r.Lines.Select(l => (l.Account, l.DebitCredit, l.Ledger, l.PostLine)).ToList());
        row["PayCode"] = "SJ";
        Assert.Equal(("7520", "MS"), (PostPayClnService.Lines([row], "C").Lines[0].Account, PostPayClnService.Lines([row], "C").Lines[0].Ledger));
    }
}

/// <summary>Payment Posting > Clinic against the real databases (skipped without CIS_TEST_DB); the post is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PostPayClnDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_post()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var service = new PostPayClnService(TestDb.Sessions(), new PostPayClnRepository(), new PostInvClnGstRepository(),
            new PostInvClnService(TestDb.Sessions(), new PostInvClnRepository(), exporter));
        var r = await service.SearchAsync(new PostPayClnRequest { PvDateFrom = "02/07/2021", PvDateTo = "02/07/2021", PvFrom = "CP100465721", PvTo = "CP100465721" });
        Assert.Contains(r.Lines, l => l.Account == "84B294" && l.Amount == 34.20m);
        var posted = await service.SearchAsync(new PostPayClnRequest { Posted = "1 - Yes", BatchFrom = "2647", BatchTo = "2647" });
        Assert.True(posted.Posted);
        Assert.Equal(PostPayClnService.NoPvDate, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PostAsync(
            new PostPayClnPostRequest(new PostPayClnRequest(), "07/2021", [new("CP100465721", "4B294")]), "DEVADMIN"))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        var batch = await new PostInvClnGstRepository().NextBatchAsync(tx);
        await new PostPayClnRepository().InsertAsync(tx, new PostPayClnTick("CP100465721", "4B294"), "07", "2021", batch, "DEVADMIN");
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISWorksheet.dbo.SUNPosting WHERE BatchNo = @batch AND TransRef = 'CP100465721' AND PostType = 'Pay2Cln'", new { batch }, tx.Sql().Transaction));
    }
}
