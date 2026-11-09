using Cis.Application.Payment;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Payment;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>FFS Payment against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PayFfsDatabaseTests
{
    [Fact]
    public async Task Search_checks_voucher_and_rolled_back_save()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new PayFfsService(TestDb.Sessions(), new PayFfsRepository(), new PayClinicRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()));
        Assert.Equal(PayFfsService.NoBordx, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new PayFfsSearchRequest()))).Message);
        var unpaid = Assert.Single((await service.SearchAsync(new PayFfsSearchRequest { BordxFrom = "KDC55955" })).Rows);
        Assert.Equal("", unpaid["Status"]);
        var paid = await service.SearchAsync(new PayFfsSearchRequest { Scan = false }, "CP100175200");
        Assert.Contains(paid.Rows, x => (string)x["Status"]! == "Full");
        Assert.Equal(PayFfsService.BadPvDate,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new PayFfsSaveRequest { Ticked = ["KDC55955"] }, "DEVADMIN"))).Message);
        Assert.EndsWith(".xlsx", (await service.VoucherAsync("CP100175200", null, "DEVADMIN")).FileName);

        await using var tx = await TestDb.Sessions().BeginAsync();
        await new PayFfsRepository().SaveAsync(tx, new PayFfsRow(null, "KDC55955", "CPTESTF", DateTime.Today, null, null, 1m, "MB", null, null, 1m, "11", "2026",
            "HEALTH CONNECT SDN BHD", "DEVADMIN"));
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISPayment.dbo.PaymentFFS WHERE PVNo = 'CPTESTF'", transaction: tx.Sql().Transaction));
    }
}
