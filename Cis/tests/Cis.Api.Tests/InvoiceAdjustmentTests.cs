using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Invoice;
using Cis.Infrastructure.Membership;
using Dapper;

namespace Cis.Api.Tests;

public class InvoiceAdjustmentRulesTests
{
    [Fact]
    public void Pending_reason_history_is_stamped_only_when_the_reasons_change()
    {
        var now = new DateTime(2026, 11, 2, 10, 30, 0);
        var loaded = new AdjustmentCase { PendingReasonIds = ["15", "16"] };
        Assert.Equal("", InvoiceAdjustmentService.RepAction(new AdjustmentExtra { Pending = "Y", PendingReasonIds = ["15", "16"] }, loaded, now));
        Assert.Equal("|02/11/2026 10:30:00-15,17", InvoiceAdjustmentService.RepAction(new AdjustmentExtra { Pending = "Y", PendingReasonIds = ["15", "17"] }, loaded, now));
        Assert.Equal("|02/11/2026 10:30:00-", InvoiceAdjustmentService.RepAction(new AdjustmentExtra { Pending = "N", PendingReasonIds = ["15"] }, loaded, now));
        Assert.Equal("", InvoiceAdjustmentService.RepAction(new AdjustmentExtra { Pending = "N" }, new AdjustmentCase(), now));
    }
}

/// <summary>Adjustment against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class InvoiceAdjustmentDatabaseTests
{
    private const string Case = "EA12586874";

    private static (InvoiceAdjustmentService Service, InvoiceAdjustmentRepository Repo) Create()
    {
        var clock = new FixedClock(new DateTime(2026, 11, 2));
        var clinicRepo = new ClinicInvoiceRepository();
        var clinic = new ClinicInvoiceService(TestDb.Sessions(), clinicRepo, new MemberSuppRepository(), new MbmAdjustmentRepository(), new MbmEnquiryRepository(),
            new WorksheetRepository(), clock);
        var repo = new InvoiceAdjustmentRepository();
        return (new InvoiceAdjustmentService(TestDb.Sessions(), repo, clinicRepo, clinic, new WorksheetRepository(), clock), repo);
    }

    private static AdjustmentRequest Request(string bill = "233", string remark = "ZZ ADJ TEST") => new(Case, new ClinicInvoiceEntry
    {
        MbmNumber = "BSHA0449584*15", CoverId = "00", ClinicCode = "T301", VisitDate = "02/08/2021", VisitTime = "10:50", Diagnoses = ["202", "101", "", "", ""],
        B = [null, 30m, 0m, 0m, 193m, null, null, null, null], A = Enumerable.Range(0, 34).Select(i => i == 18 ? (decimal?)10 : null).ToArray(),
        POthDesc = "GM", ClinicBillAmount = bill, TotalInvoiceAmount = "233", Ffs = "Y", Remark = remark, McDays = 0,
    }, new AdjustmentExtra { Status = "O", Claimability = "A", Pending = "N", PayTo = "C" });

    [Fact]
    public async Task Load_check_and_a_rolled_back_save()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var (service, repo) = Create();
        var loaded = await service.LoadAsync(Case.ToLowerInvariant());
        Assert.Equal(AdjustmentEditMode.Full, loaded.EditMode);
        Assert.Equal("T301", ((string)loaded.Case["ClinicCode"]!).Trim());
        Assert.NotEmpty(loaded.PendingReasons);
        Assert.Equal(31, loaded.Worksheet.Rows.Count);
        await Assert.ThrowsAsync<NotFoundException>(() => service.LoadAsync("ZZNOCASE1"));

        var wrongBill = await Assert.ThrowsAsync<BusinessRuleException>(() => service.UpdateAsync(Request(bill: "200"), "devadmin"));
        Assert.Equal("Please check Bill Amount! ", wrongBill.Message);
        var ws = await service.ComputeAsync(Request(), "devadmin");
        Assert.Equal(233m, ws.TotalActual);

        await using var s = await TestDb.Sessions().BeginAsync();
        var member = loaded.Member!;
        var a = ClinicInvoiceRules.Amounts(Request().Entry);
        Assert.Equal((223m, 233m), (a[1], ClinicInvoiceRules.Total(a)));
        await repo.SaveAsync(s, new AdjustmentSave(Request(), loaded with { Member = member }, a, 233, 0, ws, "", "DEVADMIN", new DateTime(2026, 11, 2), 23.3m, 233));
        var cas = await s.Sql().Connection.QuerySingleAsync<(string Remark, string User)>(
            "SELECT RTRIM(CasRemark), RTRIM(LastUpdateBy) FROM dbo.CISCAS WHERE CASNumber = @Case", new { Case }, s.Sql().Transaction);
        Assert.Equal(("ZZ ADJ TEST", "DEVADMIN"), cas);
        var claim = await s.Sql().Connection.QuerySingleAsync<(decimal Total, decimal Mco)>(
            "SELECT CLMTotalActual, CLMMCOProcFee FROM CISWorksheet.dbo.Claim WHERE CASNumber = @Case", new { Case }, s.Sql().Transaction);
        Assert.Equal((233m, 23.3m), claim);
    }
}
