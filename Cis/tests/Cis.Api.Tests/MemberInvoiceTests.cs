using Cis.Application.Invoice;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Invoice;
using Cis.Infrastructure.Membership;
using Dapper;

namespace Cis.Api.Tests;

public class MemberInvoiceRulesTests
{
    private static readonly ClinicInvoiceEntry Ok = new()
    {
        MbmNumber = "M", ClinicCode = "T301", InvoiceNo = "I1", InvoiceDate = "31/08/2021", ReceiptDate = "01/09/2021", ReceiptMode = "1",
        TotalInvoiceAmount = "45", VisitDate = "15/08/2021",
    };

    [Fact]
    public void ChkData_messages_in_desktop_order()
    {
        var x = new MemberInvoiceExtra { MemoClaim = "Y" };
        Assert.Null(MemberInvoiceService.Check(Ok, x, 45, false));
        Assert.Equal("Please Enter Invoice Receipt Mode ! ", MemberInvoiceService.Check(Ok with { ReceiptMode = "" }, x, 45, false));
        Assert.Equal("Total Invoice Amount cannot be Zero ! ", MemberInvoiceService.Check(Ok, x, 0, false));
        Assert.Equal("Please keyin the NP Clinic ! ", MemberInvoiceService.Check(Ok with { ClinicCode = "np" }, x, 45, false));
        Assert.Equal("Please Enter Accepted status ", MemberInvoiceService.Check(Ok, new MemberInvoiceExtra(), 45, false));
        Assert.Equal("Please Enter Repudiation Reason", MemberInvoiceService.Check(Ok, new MemberInvoiceExtra { MemoClaim = "N" }, 45, false));
        Assert.Equal("Please enter status for referral letter",
            MemberInvoiceService.Check(Ok with { ClinicCode = "NPS" }, new MemberInvoiceExtra { MemoClaim = "Y", NpClinic = "X" }, 45, false));
    }

    [Fact]
    public void Co_pay_rules_for_MM_and_IBM_members()
    {
        decimal[] A() { var a = new decimal[34]; a[1] = 100; return a; }
        Assert.Equal(-20m, MemberInvoiceService.CoPay(A(), new MemberInvoiceExtra { RefLetter = "N", Pharmacy = "N" }, 0.8m, false, null));
        Assert.Equal(0m, MemberInvoiceService.CoPay(A(), new MemberInvoiceExtra { RefLetter = "Y" }, 0.8m, false, null));
        var v = new DateTime(2021, 8, 15);
        Assert.Equal(-10m, MemberInvoiceService.CoPay(A(), new MemberInvoiceExtra { Pharmacy = "Y" }, 1, true, v));
        Assert.Equal(-5m, MemberInvoiceService.CoPay(A(), new MemberInvoiceExtra { ClinicType = "G", MemoClaim = "Y" }, 1, true, v));
        Assert.Equal(-25m, MemberInvoiceService.CoPay(A(), new MemberInvoiceExtra { ClinicType = "S" }, 1, true, v));
    }
}

/// <summary>Member invoice against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class MemberInvoiceDatabaseTests
{
    [Fact]
    public async Task Add_verify_and_a_rolled_back_save()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var clock = new FixedClock(new DateTime(2026, 11, 2));
        var clinicRepo = new ClinicInvoiceRepository();
        var clinic = new ClinicInvoiceService(TestDb.Sessions(), clinicRepo, new MemberSuppRepository(), new MbmAdjustmentRepository(), new MbmEnquiryRepository(),
            new WorksheetRepository(), clock);
        var repo = new MemberInvoiceRepository();
        var service = new MemberInvoiceService(TestDb.Sessions(), repo, clinicRepo, clinic, new WorksheetRepository(), clock);
        var mode = (await clinic.SetupAsync()).ReceiptModes[0].Code;
        var entry = new ClinicInvoiceEntry
        {
            InvoiceNo = "ZZTEST-MBM", InvoiceDate = "31/08/2021", ReceiptDate = "01/09/2021", ReceiptMode = mode, ClinicCode = "T301", TotalInvoiceAmount = "45",
            MbmNumber = "DCHC9086608", CoverId = "00", VisitDate = "15/08/2021", BillNo = "ZZB", Diagnoses = ["101", "", "", "", ""],
            B = [null, 30m, null, null, null, null, null, null, null], A = Enumerable.Range(0, 34).Select(i => i == 10 ? (decimal?)15 : null).ToArray(),
        };
        var request = new MemberInvoiceRequest(entry, new MemberInvoiceExtra { MemoClaim = "Y", ClinicType = "G" });
        var line = await service.AddAsync(request, "devadmin");
        Assert.Equal(45m, line.Total);
        Assert.Equal(45m, line.Worksheet.TotalActual);
        var verify = await service.VerifyAsync(request, "devadmin");
        Assert.Contains("45 Days LATE ! ", verify.Errors);

        await using var s = await TestDb.Sessions().BeginAsync();
        var caseNo = await clinicRepo.NextCaseNoAsync(s);
        var batch = await repo.NextBatchAsync(s);
        await repo.SaveAsync(s, new MemberCaseSave(caseNo, batch, request, line, "DEVADMIN", new DateTime(2026, 11, 2)));
        var cas = await s.Sql().Connection.QuerySingleAsync<(string Batch, string Clinic)>(
            "SELECT RTRIM(OPBatchNo), RTRIM(ClinicCode) FROM dbo.CISCAS WHERE CASNumber = @caseNo", new { caseNo }, s.Sql().Transaction);
        Assert.Equal((batch, "T301"), cas);
        var inv = await s.Sql().Connection.QuerySingleAsync<(string PayTo, decimal Amount)>(
            "SELECT RTRIM(INVPayTo), INVAmount FROM CISWorksheet.dbo.ClaimInvoice WHERE CASNumber = @caseNo", new { caseNo }, s.Sql().Transaction);
        Assert.Equal(("M", 45m), inv);
        Assert.Equal(1, await s.Sql().Connection.ExecuteScalarAsync<int>(
            "SELECT COUNT(*) FROM CISWorksheet.dbo.Claim WHERE CASNumber = @caseNo", new { caseNo }, s.Sql().Transaction));
    }
}
