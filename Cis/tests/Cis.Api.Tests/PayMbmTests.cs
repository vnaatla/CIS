using Cis.Application.Payment;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Payment;
using Dapper;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

/// <summary>Payment To Member rules (no database).</summary>
public class PayMbmRuleTests
{
    [Fact]
    public void Wording_lines_and_amounts_follow_the_desktop()
    {
        var (one, two) = PayMbmService.Lines(7254.15m);
        // The 50-character check is made before the joining space is added, as the desktop (line 1 can reach 51 characters).
        Assert.Equal("Seven Thousand Two Hundred And Fifty Four And Cents", one);
        Assert.Equal("Fifteen Only", two);
        var unpaid = new Dictionary<string, object?> { ["PVPayMBMStatus"] = null, ["CLMPayableByMedix2M"] = 88m, ["MBMPaidAmt"] = null, ["CLMPostageFee"] = 3m };
        PayMbmService.Shape(unpaid, 0, "");
        Assert.Equal(85m, unpaid["AmountToPay"]);
        var paid = new Dictionary<string, object?> { ["PVPayMBMStatus"] = "F", ["CLMPayableByMedix2M"] = 88m, ["MBMPaidAmt"] = 88m };
        PayMbmService.Shape(paid, 88, "F");
        Assert.Equal(PayMbmService.Paid, paid["TickError"]);
    }
}

/// <summary>Payment To Member against the real databases (skipped without CIS_TEST_DB); writes are rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class PayMbmDatabaseTests
{
    [Fact]
    public async Task Search_checks_voucher_and_rolled_back_save()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = new PayMbmService(TestDb.Sessions(), new PayMbmRepository(), new PayClinicRepository(),
            new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions()))), new PdfReportWriter()));
        Assert.Equal(PayMbmService.NoWorksheet, (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SearchAsync(new PayMbmSearchRequest { Scan = true }))).Message);
        var rows = await service.SearchAsync(new PayMbmSearchRequest { PvFrom = "CP100471601", PvTo = "CP100471603", Bordx = null });
        Assert.Equal(7, rows.Count);
        var unpaid = Assert.Single(await service.SearchAsync(new PayMbmSearchRequest { Scan = true, CaseFrom = "SLK004743875" }));
        Assert.Equal("CHOOK SHENG CAI", unpaid["PatientName"]);
        Assert.Equal(PayMbmService.NoCheque,
            (await Assert.ThrowsAsync<BusinessRuleException>(() => service.SaveAsync(new PayMbmSaveRequest { Ticked = [new("SLK004743875", null)] }, "DEVADMIN"))).Message);
        Assert.EndsWith(".xlsx", (await service.VoucherAsync(null, "CP100471601", "CP100471603", null, null, "DEVADMIN")).FileName);

        var repo = new PayMbmRepository();
        await using var tx = await TestDb.Sessions().BeginAsync();
        await repo.SaveAsync(tx, new PayMbmRow(null, "SLK004743875", "CPTESTM", DateTime.Today, DateTime.Today, "CHQ1", "X", "MB", "", null, 88m, "DEVADMIN", "11", "2026", "F"));
        await repo.PvChequeAsync(tx, "CPTESTM", DateTime.Today, "CHQ1", 88m, "X", "Eighty Eight Only", "");
        Assert.Equal("Eighty Eight Only", await tx.Sql().Connection.QuerySingleAsync<string>(
            "SELECT PVWordingOne FROM CISPayment.dbo.Payment2MBM WHERE PVNo = 'CPTESTM'", transaction: tx.Sql().Transaction));
    }
}
