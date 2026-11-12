using Cis.Application.AcExport;
using Cis.Domain.Errors;
using Cis.Infrastructure.AcExport;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Bordx Posting > Grp Company rules (no database).</summary>
public class PostBordxGrpRuleTests
{
    private static Dictionary<string, object?> Row() => new()
    {
        ["BordxRefNo"] = "KIS18019", ["BordxDate"] = new DateTime(2021, 7, 1), ["MBMGroupCompany"] = "ASIAFLEX PRODUCTS SDN. BHD.", ["CLMBordxAmt"] = 5077.34m,
        ["CLMMCOProcFee"] = 443.854m, ["GRPMCOStatus"] = "F", ["INVPayTo"] = "C", ["CLMPayableByMedix2C"] = 4633.486m, ["PAYCode"] = "DC", ["BORDXPrintMonth"] = "JULY",
        ["BORDXPrintYear"] = "2021", ["CoGRPCode"] = "A242",
    };

    [Fact]
    public void Gst_gives_a_65_debit_and_a_6000_credit()
    {
        var r = PostBordxGrpService.Lines([Row()], "G", true);
        Assert.Equal([("65DC", "D", 5077.34m), ("6000", "C", 5077.34m)], r.Lines.Select(l => (l.Account, l.DebitCredit, l.Amount)).ToList());
        Assert.Equal("2021/7-ASIAFLEX PRODUCTS SDN. BHD.", r.Lines[0].Description);
        Assert.True(r.Lines[1].PostLine);
    }

    [Fact]
    public void The_hidden_screen_adds_the_mco_fee_line_for_fee_for_service_companies()
    {
        var r = PostBordxGrpService.Lines([Row()], "G", false);
        Assert.Equal([("65DC", 5077.34m), ("6000", 4633.49m), ("1002", 443.85m)], r.Lines.Select(l => (l.Account, l.Amount)).ToList());
        Assert.Equal("2107-ASIAFLEX PRODUCTS SD", r.Lines[0].Description);
    }
}

/// <summary>Bordx Posting > Grp Company against the real databases (skipped without CIS_TEST_DB); the post is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PostBordxGrpDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_post()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var service = new PostBordxGrpService(TestDb.Sessions(), new PostBordxGrpRepository(), new PostInvClnGstRepository(),
            new PostInvClnService(TestDb.Sessions(), new PostInvClnRepository(), exporter));
        // GST swaps the two bordx no. boxes, so the same no. in both still finds it.
        var r = await service.SearchAsync(new PostBordxGrpRequest { BordxFrom = "KIS18019", BordxTo = "KIS18019" }, true);
        Assert.Contains(r.Lines, l => l.Reference == "KIS18019" && l.Account == "6000" && l.Amount == 5077.34m);
        Assert.Equal(PostInvClnService.NoBordxDate, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PostAsync(
            new PostBordxGrpPostRequest(new PostBordxGrpRequest(), "07/2021", ["KIS18019"]), "DEVADMIN", true))).Message);

        await using var tx = await TestDb.Sessions().BeginAsync();
        var batch = await new PostInvClnGstRepository().NextBatchAsync(tx);
        await new PostBordxGrpRepository().InsertAsync(tx, "KIS18019", "07", "2021", batch, "DEVADMIN");
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISWorksheet.dbo.SUNPosting WHERE BatchNo = @batch AND TransRef = 'KIS18019' AND PostType = 'BordxGrp'", new { batch }, tx.Sql().Transaction));
    }
}
