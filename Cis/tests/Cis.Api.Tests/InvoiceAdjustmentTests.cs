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

    /// <summary>The form of a loaded case (as the page fills it).</summary>
    private static AdjustmentRequest FromCase(AdjustmentCase c)
    {
        string S(string k) => InvoiceAdjustmentService.T(c.Case, k);
        decimal? N(string k) => ClinicInvoiceRules.NumOrNull(S(k));
        string[] bCols = ["", "GPFee", "GPLTMed", "GPInj", "GPMed", "SPFee", "SPLTMed", "SPInj", "SPMed"];
        var aCols = new Dictionary<int, string>
        {
            [7] = "CLMDental", [8] = "CLMMaternity", [9] = "CLMPRVCheckup", [10] = "CLMXRay", [11] = "CLMUltraSound", [12] = "CLMECG", [13] = "CLMDressing",
            [14] = "CLMNebulizer", [15] = "CLMID", [16] = "CLMTSSTO", [17] = "CLMMinorSurgery", [18] = "CLMPOthers", [19] = "CLMBloodDrawing", [20] = "CLMBloodGlucose",
            [21] = "CLMUrine", [22] = "CLMPapSmear", [23] = "CLMLOthers", [24] = "CLMDTPPolio", [25] = "CLMDTPPolioHIB", [26] = "CLMHepatitisB", [27] = "CLMMMR",
            [28] = "CLMIOthers", [30] = "CLMVision", [32] = "CLMTyphiadJabNoDed", [33] = "CLMOthersNoDed",
        };
        var b = Enumerable.Range(0, 9).Select(i => i == 0 ? null : N(bCols[i])).ToArray();
        var a = Enumerable.Range(0, 34).Select(i => aCols.TryGetValue(i, out var col) ? N(col) : null).ToArray();
        var entry = new ClinicInvoiceEntry
        {
            MbmNumber = S("MBMnumber"), CoverId = S("CoverID"), ClinicCode = S("ClinicCode"),
            VisitDate = DateTime.Parse(S("DateVisit"), System.Globalization.CultureInfo.InvariantCulture).ToString("dd/MM/yyyy", System.Globalization.CultureInfo.InvariantCulture),
            Diagnoses = [S("CASTreatReason1"), S("CASTreatReason2"), S("CASTreatReason3"), S("CASTreatReason4"), S("CASTreatReason5")], B = b, A = a, Ffs = "N",
        };
        var total = ClinicInvoiceRules.Total(ClinicInvoiceRules.Amounts(entry));
        return new AdjustmentRequest(S("CASNumber"), entry with { ClinicBillAmount = total.ToString(System.Globalization.CultureInfo.InvariantCulture) },
            new AdjustmentExtra { Status = S("CASStatus"), Claimability = S("Claimability"), Pending = S("Pending") is { Length: > 0 } p ? p : "N", PayTo = S("INVPayTo") });
    }

    [Fact]
    public async Task After_posting_creates_a_sub_case_only_for_a_posted_case()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var (service, repo) = Create();
        var notPosted = await service.LoadAfterPostAsync(Case);
        Assert.Equal((AdjustmentEditMode.None, InvoiceAdjustmentService.NotPosted), (notPosted.EditMode, notPosted.EditMessage));
        var loaded = await service.LoadAfterPostAsync("EA12549445");
        Assert.Equal(AdjustmentEditMode.Full, loaded.EditMode);
        var request = FromCase(loaded);
        var wrongBill = await Assert.ThrowsAsync<BusinessRuleException>(() =>
            service.CreateAfterPostAsync(request with { Entry = request.Entry with { ClinicBillAmount = "1" } }, "devadmin"));
        Assert.Equal("Please check Bill Amount! ", wrongBill.Message);

        await using var s = await TestDb.Sessions().BeginAsync();
        var ws = await service.ComputeAsync(request, "devadmin");
        var a = ClinicInvoiceRules.Amounts(request.Entry);
        var total = ClinicInvoiceRules.Total(a);
        var newCase = await new ClinicInvoiceRepository().NextCaseNoAsync(s);
        await repo.SaveAsync(s, new AdjustmentSave(request, loaded, a, total, 0, ws, "", "DEVADMIN", new DateTime(2026, 11, 2), 0, total, NewCase: newCase));
        var cas = await s.Sql().Connection.QuerySingleAsync<(string Master, string Clinic)>(
            "SELECT RTRIM(MasterCASNumber), RTRIM(ClinicCode) FROM dbo.CISCAS WHERE CASNumber = @newCase", new { newCase }, s.Sql().Transaction);
        Assert.Equal(("EA12549445", "N197"), cas);
        Assert.Equal(1, await s.Sql().Connection.ExecuteScalarAsync<int>(
            "SELECT COUNT(*) FROM CISWorksheet.dbo.Claim WHERE CASNumber = @newCase", new { newCase }, s.Sql().Transaction));
    }

    [Fact]
    public async Task After_bordx_saves_only_the_case_and_the_invoice()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var (service, repo) = Create();
        var loaded = await service.LoadAfterBordxAsync(Case);
        Assert.Equal(AdjustmentEditMode.Full, loaded.EditMode);
        var wrongBill = await Assert.ThrowsAsync<BusinessRuleException>(() => service.UpdateAfterBordxAsync(Request(bill: "200"), "devadmin"));
        Assert.Equal("Please check Bill Amount! ", wrongBill.Message);

        await using var s = await TestDb.Sessions().BeginAsync();
        var before = await s.Sql().Connection.QuerySingleAsync<decimal>(
            "SELECT CLMTotalActual FROM CISWorksheet.dbo.Claim WHERE CASNumber = @Case", new { Case }, s.Sql().Transaction);
        var ws = await service.ComputeAsync(Request(), "devadmin");
        var a = ClinicInvoiceRules.Amounts(Request().Entry);
        await repo.SaveAsync(s, new AdjustmentSave(Request(remark: "ZZ AFTER BORDX"), loaded, a, 233, 0, ws, "", "DEVADMIN", new DateTime(2026, 11, 2), 0, 233,
            AfterBordx: true));
        var cas = await s.Sql().Connection.QuerySingleAsync<(string Remark, string? Note)>(
            "SELECT RTRIM(CasRemark), AddPendNote FROM dbo.CISCAS WHERE CASNumber = @Case", new { Case }, s.Sql().Transaction);
        Assert.Equal(("ZZ AFTER BORDX", null), cas);
        Assert.Equal(before, await s.Sql().Connection.QuerySingleAsync<decimal>(
            "SELECT CLMTotalActual FROM CISWorksheet.dbo.Claim WHERE CASNumber = @Case", new { Case }, s.Sql().Transaction));
    }
}
