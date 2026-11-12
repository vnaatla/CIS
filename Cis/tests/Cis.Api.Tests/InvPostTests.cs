using Cis.Application.AcExport;
using Cis.Domain.Errors;
using Cis.Infrastructure.AcExport;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Invoice Posting rules (no database).</summary>
public class InvPostRuleTests
{
    [Fact]
    public void Summary_credit_and_detail_debit_lines()
    {
        var reg = new DateTime(2021, 8, 1);
        var summary = new Dictionary<string, object?> { ["InvoiceNo"] = "INV1", ["ClinicCode"] = "A010", ["RegisterDate"] = reg, ["CLMPayableByMedix2C"] = 49.5m };
        var dc = new Dictionary<string, object?>
        {
            ["SumInvoiceNo"] = "INV1", ["SumRegisterDate"] = reg, ["InvoiceNo"] = "INV1", ["PAYCode"] = "DC", ["GRPCoCode"] = "E020", ["CLMPayableByMedix2C"] = 49.5m,
        };
        var r = InvPostService.Lines([summary], [dc]);
        Assert.Equal([("84A010", "C"), ("60E020", "D")], r.Lines.Select(l => (l.Account, l.DebitCredit)).ToList());
        Assert.Equal((49.5m, 49.5m), (r.Debit, r.Credit));
    }
}

/// <summary>Invoice Posting against the real databases (skipped without CIS_TEST_DB); the mark is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class InvPostDatabaseTests
{
    [Fact]
    public async Task Search_checks_and_rolled_back_mark()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var service = new InvPostService(TestDb.Sessions(), new InvPostRepository(), new PostInvClnService(TestDb.Sessions(), new PostInvClnRepository(), exporter),
            TimeProvider.System);
        Assert.Equal(InvPostService.FutureDate,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new InvPostRequest(null, null, null, "01/01/2999", null, "0 - No")))).Message);
        var r = await service.SearchAsync(new InvPostRequest(null, null, null, "01/08/2021", "02/08/2021", "0 - No"));
        Assert.True(r.Records > 0);
        Assert.Contains(r.Lines, l => l.Account == "84A010");

        await using var tx = await TestDb.Sessions().BeginAsync();
        var row = (await new InvPostRepository().RowsAsync(tx, null, null, null, new DateTime(2021, 8, 1), new DateTime(2021, 8, 2), "0"))
            .First(x => x["InvoiceNo"] is string { Length: > 0 });
        var (inv, cln) = (((string)row["InvoiceNo"]!).Trim(), ((string)row["ClinicCode"]!).Trim());
        await new InvPostRepository().MarkAsync(tx, inv, cln, "DEVADMIN");
        Assert.Equal(0, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISDB.dbo.CISCas WHERE InvoiceNo = @inv AND ClinicCode = @cln AND CASInvPost = '0'", new { inv, cln }, tx.Sql().Transaction));
    }
}
