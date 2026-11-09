using Cis.Application.Payment;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Payment;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Payment To Clinic list rules (no database).</summary>
public class PayClinicRuleTests
{
    [Fact]
    public void Shape_follows_the_desktop_list()
    {
        var unpaid = new Dictionary<string, object?> { ["PVPayCLNStatus"] = null, ["CLMPayableByMedix2C"] = 40m, ["CLMNonGPFee"] = 5m, ["CLNPaidAmt"] = 0m };
        PayClinicService.Shape(unpaid);
        Assert.Equal(45m, unpaid["Balance"]);
        var partial = new Dictionary<string, object?> { ["PVPayCLNStatus"] = "P", ["CLMPayableByMedix2C"] = 40m, ["CLMNonGPFee"] = 5m, ["CLNPaidAmt"] = 10m };
        PayClinicService.Shape(partial);
        Assert.Equal("Partial", partial["Status"]);
        Assert.Equal(30m, partial["Balance"]);
    }
}

/// <summary>Payment To Clinic against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PayClinicDatabaseTests
{
    [Fact]
    public async Task Search_checks_rolled_back_save_and_voucher()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new PayClinicService(TestDb.Sessions(), new PayClinicRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()));
        Assert.Equal(PayClinicService.NoWorksheet,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new PayClinicSearchRequest { Scan = true }))).Message);
        var rows = await service.SearchAsync(new PayClinicSearchRequest { InvFrom = "01/08/2021", InvTo = "05/08/2021", Payment = 2 });
        Assert.Contains(rows, x => ((string)x["CASNumber"]!).Trim() == "EA12565564");
        Assert.Equal(PayClinicService.BadPvDate, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.AddAsync(new PayClinicSaveRequest(), "DEVADMIN"))).Message);
        Assert.Equal(PayClinicService.NoBank, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.AddAsync(new PayClinicSaveRequest
        {
            PvDate = "01/10/2021", ChequeAmt = "1", Remark = "test", Ticked = [new("EA12565564", null)],
        }, "DEVADMIN"))).Message);

        // Before the transaction below: its uncommitted Payment2CLN row would block this read.
        var file = await service.VoucherAsync("CP100472166", null, "DEVADMIN");
        Assert.EndsWith(".xlsx", file.FileName);

        var repo = new PayClinicRepository();
        await using var tx = await TestDb.Sessions().BeginAsync();
        var pv = await repo.NextPvAsync(tx, "MB", "CP");
        Assert.StartsWith("CP", pv);
        await repo.SaveAsync(tx, new PayClinicRow(null, false, "EA12565564", pv, new DateTime(2021, 10, 1), null, null, 10m, "MB", null, null, 10m, "10", "2021", "F",
            "DEVADMIN"));
        Assert.Equal(1, await tx.Sql().Connection.QuerySingleAsync<int>(
            "SELECT COUNT(*) FROM CISPayment.dbo.Payment2CLN WHERE PVCLNNo = @pv", new { pv }, tx.Sql().Transaction));
    }
}
