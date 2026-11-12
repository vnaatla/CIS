using Cis.Application.AcExport;
using Cis.Domain.Errors;
using Cis.Infrastructure.AcExport;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Receipt Posting > Grp Company rules (no database).</summary>
public class PostRecGrpRuleTests
{
    [Fact]
    public void A_receipt_debits_the_bank_and_credits_the_company_account_once()
    {
        var row = new Dictionary<string, object?>
        {
            ["ReceiptNo"] = "RP37386", ["BordxRefNo"] = "KIS1", ["PAYAccCode"] = "63A011", ["GRPCoCode"] = "P262", ["CheckAmt"] = 103547.24m, ["PAYCode"] = "AN",
        };
        var r = PostRecGrpService.Lines([row], "G");
        Assert.Equal([("7610", "D", false), ("65A011", "C", true)], r.Lines.Select(l => (l.Account, l.DebitCredit, l.PostLine)).ToList());
        var second = new Dictionary<string, object?>(row) { ["BordxRefNo"] = "KIS2" };
        var twice = PostRecGrpService.Lines([row, second], "G");
        Assert.Equal((2, 1, 103547.24m), (twice.Lines.Count, twice.Records, twice.Credit));
        Assert.Equal("KIS1, KIS2", twice.Lines[0].Extra!["Bordx"]);
        row["BordxRefNo"] = "KDC54665";
        Assert.Equal("65P262", PostRecGrpService.Lines([row], "G").Lines[1].Account);
        Assert.Equal(("7610", "D"), (PostRecGrpService.Lines([row], "D").Lines[0].Account, PostRecGrpService.Lines([row], "D").Lines[0].DebitCredit));
    }
}

/// <summary>Receipt Posting > Grp Company against the real databases (skipped without CIS_TEST_DB); the post is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PostRecGrpDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_post()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var service = new PostRecGrpService(TestDb.Sessions(), new PostRecGrpRepository(), new PostInvClnGstRepository(),
            new PostInvClnService(TestDb.Sessions(), new PostInvClnRepository(), exporter));
        var r = await service.SearchAsync(new PostRecGrpRequest { ReceiptFrom = "RP37386", ReceiptTo = "RP37386" });
        Assert.Equal(1, r.Records);
        Assert.Equal(2, r.Lines.Count);
        Assert.Equal(103547.24m, r.Credit);
        Assert.Equal(PostRecGrpService.NoReceiptDate, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PostAsync(
            new PostRecGrpPostRequest(new PostRecGrpRequest(), "07/2021", ["RP37386"]), "DEVADMIN"))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        var batch = await new PostInvClnGstRepository().NextBatchAsync(tx);
        await new PostRecGrpRepository().InsertAsync(tx, "RP37386", "07", "2021", batch, "DEVADMIN");
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISWorksheet.dbo.SUNPosting WHERE BatchNo = @batch AND TransRef = 'RP37386' AND PostType = 'RecGrp'", new { batch }, tx.Sql().Transaction));
    }
}
