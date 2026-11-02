using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Invoice;
using Dapper;

namespace Cis.Api.Tests;

public class BatchTransferServiceTests
{
    [Fact]
    public async Task Checks_use_the_desktop_messages_in_order()
    {
        var service = new BatchTransferService(new FakeSessions(), null!, new FixedClock(new DateTime(2026, 11, 2)));
        Assert.Equal(BatchTransferService.Invalid, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.BatchAsync(" "))).Message);
        Assert.Equal(BatchTransferService.BatchFirst, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SaveAsync(new BatchTransferRequest("OPY9978", " "), "DEV"))).Message);
    }

    [Fact]
    public void ShowInfo_joins_invoices_and_rounds_each_amount()
    {
        var one = BatchTransferService.Info("B1", [new BatchRow { TotNoOfCase = 4, InvoiceNo = "INV1 ", ClnName = "KLINIK A", InvoiceStmtAmt = 184m }]);
        Assert.Equal(new BatchInfo("B1", 4, "INV1", 184m, "KLINIK A", null), one);
        var two = BatchTransferService.Info("B2", [
            new BatchRow { TotNoOfCase = 1, InvoiceNo = "INV1", ClnName = "A", InvoiceStmtAmt = 10.005m },
            new BatchRow { TotNoOfCase = 2, InvoiceNo = "INV1", ClnName = "B", InvoiceStmtAmt = 5m },
            new BatchRow { TotNoOfCase = 3, InvoiceNo = "INV2", ClnName = "C", InvoiceStmtAmt = 1m },
        ]);
        Assert.Equal(("INV1,INV2", 6, 16.01m, "C"), (two.InvoiceNo, two.Cases, two.InvoiceAmt, two.Clinic));
        Assert.Equal("Invoice No : INV1,INV2, Different Clinic Found !", two.Warning);
        Assert.Equal(BatchTransferService.Invalid, Assert.Throws<NotFoundException>(() => BatchTransferService.Info("X", [])).Message);
    }
}

/// <summary>Batch Transfer against the real CISDB (skipped without CIS_TEST_DB); the transfer is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class BatchTransferDatabaseTests
{
    [Fact]
    public async Task Batches_show_their_totals_and_the_cases_move()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new BatchTransferService(TestDb.Sessions(), new BatchTransferRepository(), new FixedClock(new DateTime(2026, 11, 2)));
        var a = await service.BatchAsync("opy9978");
        Assert.Equal(new BatchInfo("OPY9978", 4, "H153/SEPT14 - SYASYA", 184m, "KLINIK DAN SURGERI HILLTOP", null), a);
        await Assert.ThrowsAsync<NotFoundException>(() => service.BatchAsync("ZZNOBATCH"));
        Assert.Equal(BatchTransferService.DifferentInvoice, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SaveAsync(new BatchTransferRequest("OPY9978", "OPY9987"), "DEV"))).Message);
        Assert.Equal(BatchTransferService.SameBatch, (await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.SaveAsync(new BatchTransferRequest("OPY9978", "opy9978"), "DEV"))).Message);

        var repo = new BatchTransferRepository();
        await using var s = await TestDb.Sessions().BeginAsync();
        Assert.Equal(4, await repo.SaveAsync(s, "OPY9978", "OPY9987", "Transfer By : DEV OPY9978->OPY9987 , 20261102", 283m));
        var rows = (await s.Sql().Connection.QueryAsync<(string Batch, decimal Amt, string? Remark)>(
            "SELECT OPBatchNo, InvoiceStmtAmt, BatchTransRemarks FROM dbo.CISCAS WHERE OPBatchNo IN ('OPY9978', 'OPY9987')", transaction: s.Sql().Transaction)).ToList();
        Assert.Equal(6, rows.Count);
        Assert.All(rows, r => Assert.Equal(("OPY9987", 283m), (r.Batch, r.Amt)));
        Assert.Equal(4, rows.Count(r => r.Remark == "Transfer By : DEV OPY9978->OPY9987 , 20261102"));
    }
}
