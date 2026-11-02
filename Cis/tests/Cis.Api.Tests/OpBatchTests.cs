using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Invoice;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

public class OpBatchRulesTests
{
    private static Dictionary<string, object?> Row(string batch, string bill, string pending, string claimability) =>
        new() { ["OPBatchNo"] = batch, ["BillNo"] = bill, ["Pending"] = pending, ["Claimability"] = claimability };

    [Fact]
    public void ItemCheck_messages_follow_the_desktop_order()
    {
        Assert.Equal(OpBatchService.AlreadyDone, OpBatchService.CheckMessage(Row("OCA1", "INV", "Y", "X")));
        Assert.Null(OpBatchService.CheckMessage(Row("OCA1", "", "N", "A")));
        Assert.Equal(OpBatchService.IsPending, OpBatchService.CheckMessage(Row("", "INV", "Y", "A")));
        Assert.Equal(OpBatchService.Claimability, OpBatchService.CheckMessage(Row("", "", "N", "")));
    }

    [Fact]
    public void Invoice_number_and_month_check()
    {
        Assert.Equal("INV9-T301", OpBatchService.InvoiceNumber("INV9", "T301 - KLINIK X"));
        Assert.Equal("INV9-T30", OpBatchService.InvoiceNumber("INV9", "T30"));
        Assert.False(OpBatchService.MultipleMonths([new DateTime(2021, 8, 1), new DateTime(2021, 8, 31)]));
        Assert.True(OpBatchService.MultipleMonths([new DateTime(2021, 8, 1), new DateTime(2021, 9, 1)]));
    }
}

/// <summary>OP Batch/Invoice against the real databases (skipped without CIS_TEST_DB); updates are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class OpBatchDatabaseTests
{
    private static OpBatchFilters T301(string? from = "01/08/2021", string? to = "31/08/2021") =>
        new(false, false, null, null, false, null, null, null, null, "T301 - X", null, null, from, to, false, null, null, null, null, null, null, false,
            null, null, null, null, null, null, null);

    [Fact]
    public async Task Search_checks_and_a_batch_with_an_invoice()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var repo = new OpBatchRepository();
        var exporter = new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter());
        var service = new OpBatchService(TestDb.Sessions(), repo, exporter, new FixedClock(new DateTime(2026, 11, 2)));
        var list = await service.SearchAsync(T301());
        Assert.Equal(["EA12585463", "EA12586874", "EA12586891"], list.Rows.Select(r => (string)r["CASNumber"]!).Order());
        Assert.False(list.MultipleMonths);
        Assert.True((await service.ExportAsync(T301())).Content.Length > 0);
        Assert.Equal(OpBatchService.VisitFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SaveAsync(new OpBatchRequest(T301(to: ""), ["EA12586874"], "INV9"), "DEV"))).Message);
        Assert.Equal(OpBatchService.SelectRecord, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SaveAsync(new OpBatchRequest(T301(), [], ""), "DEV"))).Message);

        await using var s = await TestDb.Sessions().BeginAsync();
        var batch = await repo.NextBatchAsync(s, "OC");
        Assert.Matches("^OC[A-Z][0-9]{7}$", batch);
        await repo.SaveAsync(s, "EA12586874", batch, "DEV", new DateTime(2026, 11, 2), "ZZINV-T301", new DateTime(2021, 8, 31));
        var row = await s.Sql().Connection.QuerySingleAsync<(string OPBatchNo, string OPNoEnterUser, string InvoiceNo, string CASInvMonth, string CASInvYear)>(
            "SELECT OPBatchNo, OPNoEnterUser, InvoiceNo, CASInvMonth, CASInvYear FROM dbo.CISCAS WHERE CASNumber = 'EA12586874'", transaction: s.Sql().Transaction);
        Assert.Equal((batch, "DEV", "ZZINV-T301", "08", "2021"), row);
        Assert.True(await repo.InvoiceUsedAsync(s, "ZZINV-T301"));
    }
}
