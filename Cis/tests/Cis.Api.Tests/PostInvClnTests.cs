using Cis.Application.AcExport;
using Cis.Domain.Errors;
using Cis.Infrastructure.AcExport;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Invoice Posting > Clinic rules (no database).</summary>
public class PostInvClnRuleTests
{
    [Fact]
    public void Each_invoice_gives_a_credit_and_a_debit_line()
    {
        var row = new Dictionary<string, object?>
        {
            ["InvoiceNo"] = "0210152590", ["ClinicCode"] = "P320", ["ClnName"] = "KLINIK X", ["InvoiceDate"] = new DateTime(2021, 7, 1), ["CLMPayableByMedix2C"] = 983.004m,
        };
        var r = PostInvClnService.Lines([row], "G");
        Assert.Equal(2, r.Lines.Count);
        Assert.Equal(("84P320", "C", false), (r.Lines[0].Account, r.Lines[0].DebitCredit, r.Lines[0].PostLine));
        Assert.Equal(("6000", "D", "2107-KLINIK X", true), (r.Lines[1].Account, r.Lines[1].DebitCredit, r.Lines[1].Description, r.Lines[1].PostLine));
        Assert.Equal(983.00m, r.Debit);
        Assert.Single(PostInvClnService.Lines([row], "D").Lines);
    }
}

/// <summary>Invoice Posting > Clinic against the real databases (skipped without CIS_TEST_DB); the post is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PostInvClnDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_post()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new PostInvClnService(TestDb.Sessions(), new PostInvClnRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()));
        Assert.NotEmpty(await service.PeriodsAsync());
        var r = await service.SearchAsync(new PostInvClnRequest { InvDateFrom = "01/07/2021", InvDateTo = "01/07/2021" });
        Assert.Equal(116, r.Records);
        Assert.Equal(232, r.Lines.Count);
        Assert.Equal(r.Debit, r.Credit);
        var search = new PostInvClnRequest { InvDateFrom = "01/07/2021", InvDateTo = "01/07/2021" };
        Assert.Equal(PostInvClnService.NoBordxDate, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.PostAsync(
            new PostInvClnPostRequest(search, "07/2021", [new("0210152590", "P320", new DateTime(2021, 7, 1))]), "DEVADMIN"))).Message);

        var repo = new PostInvClnRepository();
        await using var tx = await TestDb.Sessions().BeginAsync();
        var batch = await repo.NextInvNoAsync(tx);
        Assert.Matches("^[A-Z][0-9]{6}$", batch);
        await repo.InsertAsync(tx, new PostInvClnTick("0210152590", "P320", new DateTime(2021, 7, 1)), "07", "2021", batch, "DEVADMIN");
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISWorksheet.dbo.SUNPosting WHERE BatchNo = @batch AND InvoiceNo = '0210152590'", new { batch }, tx.Sql().Transaction));
    }
}
