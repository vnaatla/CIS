using Cis.Application.Payment;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Payment;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>To Member (old) against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PayMbmOldDatabaseTests
{
    [Fact]
    public async Task Search_checks_voucher_and_rolled_back_save()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new PayMbmOldService(TestDb.Sessions(), new PayMbmOldRepository(), new PayMbmRepository(), new PayClinicRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()));
        var row = Assert.Single(await service.SearchAsync(new PayMbmSearchRequest { Scan = true, CaseFrom = "SLK004743875" }));
        Assert.Equal("CHOOK SHENG CAI", row["PatientName"]);
        Assert.Equal(4, (await service.SearchAsync(new PayMbmSearchRequest(), "CP100471601")).Count);
        Assert.Equal(PayMbmOldService.NoPayTo, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new PayMbmOldSaveRequest(), "DEVADMIN"))).Message);
        Assert.Equal(PayMbmOldService.NoSelection, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new PayMbmOldSaveRequest
        {
            PayTo = "X", PvDate = "01/11/2026", ChequeNo = "C1", ChequeDate = "01/11/2026", ChequeAmt = "88",
        }, "DEVADMIN"))).Message);
        Assert.EndsWith(".xlsx", (await service.VoucherAsync("CP100471601", "X", "DEVADMIN")).FileName);

        await using var tx = await TestDb.Sessions().BeginAsync();
        await new PayMbmOldRepository().SaveAsync(tx, new PayMbmOldRow(false, null, "SLK004743875", "CPTESTO", DateTime.Today, DateTime.Today, "C1", 88m, "X", "MB", null,
            null, 88m, "DEVADMIN", "11", "2026", "F", "Eighty Eight Only", ""));
        Assert.Equal(88m, await tx.Sql().Connection.QuerySingleAsync<decimal>(
            "SELECT MBMPaidAmt FROM CISPayment.dbo.Payment2MBM WHERE PVNo = 'CPTESTO'", transaction: tx.Sql().Transaction));
    }
}
