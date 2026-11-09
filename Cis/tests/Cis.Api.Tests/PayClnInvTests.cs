using Cis.Application.Payment;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Payment;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Payment To Clinic - Invoice rules (no database).</summary>
public class PayClnInvRuleTests
{
    [Fact]
    public void Cheque_numbers_and_tick_rules_follow_the_desktop()
    {
        Assert.Equal("CHQ000124", PayClnInvService.NextCheque("CHQ000123"));
        var row = new Dictionary<string, object?> { ["PVPayCLNStatus"] = null, ["CLMPayableByMedix2C"] = 90m, ["CLNPaidAmt"] = null, ["CLNPName"] = "X", ["CLNPAdd1"] = "" };
        PayClnInvService.Shape(row);
        Assert.Equal(PayClnInvService.NoPayeeAddress, row["TickError"]);
        Assert.Equal(90m, row["Balance"]);
    }
}

/// <summary>Payment To Clinic - Invoice against the real databases (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PayClnInvDatabaseTests
{
    [Fact]
    public async Task Search_popup_voucher_ranges_and_rolled_back_save()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new PayClnInvService(TestDb.Sessions(), new PayClnInvRepository(), new PayClinicRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()));
        var rows = await service.SearchAsync(new PayClnInvSearchRequest { Clinic = "G029 - X", YearFrom = "2021", YearTo = "2021", MonthFrom = "5", MonthTo = "5" });
        Assert.Contains(rows, x => (string?)x["Status"] == "" && (string)x["CLNPName"]! == "PULAU PINANG CLINIC SDN BHD");
        Assert.NotEmpty(await service.InvoicesAsync("G029", "05", "2021"));
        Assert.EndsWith(".xlsx", (await service.VoucherAsync("CP100468765", null, "DEVADMIN")).FileName);
        Assert.Contains(await service.PvRangeAsync(new PayClnInvRangeRequest("CP100468765", "CP100468765")), x => ((string)x["ClinicCode"]!).Trim() == "G029");
        Assert.Equal(PayClnInvService.BadChequeRange,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.ChequesAsync(new PayClnInvRangeRequest("CHQ000010", "CHQ000005")))).Message);
        Assert.Equal(PayClnInvService.NoCheque,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new PayClnInvSaveRequest { Ticked = [new("G029", "05", "2021")] }, "DEVADMIN"))).Message);

        var repo = new PayClnInvRepository();
        await using var tx = await TestDb.Sessions().BeginAsync();
        var cases = await repo.CasesAsync(tx, "G029", "2021", "05", 0);
        Assert.Contains(cases, c => ((string)c["CASNumber"]!).Trim() == "EA12554835");
        await repo.SaveAsync(tx, new PayClnInvRow(null, "EA12554835", "CPTEST1", DateTime.Today, DateTime.Today, "CHQ000001", 1m, "MB", null, 1m, "11", "2026", "F",
            "One Dollar Only", "", "DEVADMIN"));
        Assert.Equal("One Dollar Only", await tx.Sql().Connection.QuerySingleAsync<string>(
            "SELECT PVWordingOne FROM CISPayment.dbo.Payment2CLN WHERE PVCLNNo = 'CPTEST1'", transaction: tx.Sql().Transaction));
    }
}
