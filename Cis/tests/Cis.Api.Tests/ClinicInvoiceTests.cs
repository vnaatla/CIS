using Cis.Application.Invoice;
using Cis.Domain.Errors;
using Cis.Infrastructure.Data;
using Cis.Infrastructure.Invoice;
using Cis.Infrastructure.Membership;
using Dapper;

namespace Cis.Api.Tests;

public class ClinicInvoiceRulesTests
{
    private static decimal?[] A(params (int I, decimal V)[] v)
    {
        var a = new decimal?[34];
        foreach (var (i, x) in v) a[i] = x;
        return a;
    }

    private static decimal?[] B(params (int I, decimal V)[] v)
    {
        var b = new decimal?[9];
        foreach (var (i, x) in v) b[i] = x;
        return b;
    }

    [Fact]
    public void GetTotal_places_the_GP_and_SP_fees_and_leaves_out_the_co_pay()
    {
        var e = new ClinicInvoiceEntry { B = B((1, 20), (2, 5), (5, 30), (7, 1), (8, 1)), A = A((10, 15), (29, 4), (31, 100)) };
        var a = ClinicInvoiceRules.Amounts(e);
        Assert.Equal((0m, 0m, 25m, 0m, 32m, 0m), (a[1], a[2], a[3], a[4], a[5], a[6]));
        Assert.Equal(172m, ClinicInvoiceRules.Total(a));
        var grid = ClinicInvoiceRules.GridItems(a);
        Assert.Equal(31, grid.Length);
        Assert.Equal((25m, 15m), (grid[2]!.Value, grid[6]!.Value));
    }

    [Fact]
    public void Utilisation_and_claim_modes()
    {
        InvClaim[] claims =
        [
            new() { ClmBordxAmt = 100, ClmOthersNoDed = 10, ClmTyphiadJabNoDed = 5, ClmMaternity = 20, ClmDental = 30, ClnType = "G" },
            new() { ClmBordxAmt = 50, ClnType = "S" },
            new() { ClmBordxAmt = null, ClnType = "G" },
        ];
        Assert.Equal((135m, 0m), ClinicInvoiceRules.Utilisation(claims, false, false, false));
        Assert.Equal((35m, 50m), ClinicInvoiceRules.Utilisation(claims, true, true, true));
        Assert.Equal(0, ClinicInvoiceRules.ClaimMode("A", "02"));
        Assert.Equal((1, 2, 3), (ClinicInvoiceRules.ClaimMode("B", "00"), ClinicInvoiceRules.ClaimMode("D", "01"), ClinicInvoiceRules.ClaimMode("F", "01")));
        Assert.Null(ClinicInvoiceRules.ClaimMode("Z", "00"));
    }

    [Fact]
    public void ChkData_messages_in_desktop_order()
    {
        var today = new DateTime(2026, 11, 2);
        DateTime eff = new(2021, 8, 1), exp = new(2021, 12, 31);
        var ok = new ClinicInvoiceEntry
        {
            MbmNumber = "M", VisitDate = "15/08/2021", ClinicCode = "T301", InvoiceNo = "I1", InvoiceDate = "31/08/2021", ReceiptDate = "01/09/2021",
            TotalInvoiceAmount = "45", ClinicBillAmount = "45.00", Claimability = "A",
        };
        string? C(ClinicInvoiceEntry e, decimal total = 45) => ClinicInvoiceRules.Check(e, new decimal[34], total, eff, exp, false, today);
        Assert.Null(C(ok));
        Assert.Equal("Please Enter Visit Date! ", C(ok with { VisitDate = "" }));
        Assert.Equal("Visit Date cannot less than Membership Effective date ! ", C(ok with { VisitDate = "31/07/2021" }));
        Assert.Equal("Total Invoice Amount cannot be Zero ! ", C(ok, 0));
        Assert.Equal("Please check Bill Amount! ", C(ok with { ClinicBillAmount = "40" }));
        Assert.Equal("Please Keyin Others Diagnosis ! ", ClinicInvoiceRules.Check(ok, new decimal[34], 45, eff, exp, true, today));
        Assert.Equal("Membership has been terminated before the entered visit date.",
            ClinicInvoiceRules.VisitCheck(new DateTime(2021, 9, 1), today, true, "C", exp, "15/08/2021", "", ""));
        var errors = ClinicInvoiceRules.Verify(ok, 45, eff, exp, "30.00", "FULL", ["211"], today);
        Assert.Equal(["Diagnosis '211' is not payable ! ", "45 Days LATE ! ", "Bill Amt > Balance Limit ! "], errors);
    }
}

/// <summary>Clinic invoice against the real databases (skipped without CIS_TEST_DB); the save is rolled back.</summary>
[Collection(DatabaseCollection.Name)]
public class ClinicInvoiceDatabaseTests
{
    private static ClinicInvoiceService Service() => new(TestDb.Sessions(), new ClinicInvoiceRepository(), new MemberSuppRepository(),
        new MbmAdjustmentRepository(), new MbmEnquiryRepository(), new WorksheetRepository(), new FixedClock(new DateTime(2026, 11, 2)));

    private static ClinicInvoiceEntry Entry() => new()
    {
        InvoiceNo = "ZZTEST-INV", InvoiceDate = "31/08/2021", ReceiptDate = "01/09/2021", ReceiptMode = "", ClinicCode = "t301", TotalInvoiceAmount = "45.00",
        MbmNumber = "DCHC9086608", CoverId = "00", VisitDate = "15/08/2021", VisitTime = "10:30", BillNo = "zzb1", McDays = 1, Diagnoses = ["101", "", "", "", ""],
        B = [null, 30m, null, null, null, null, null, null, null], A = Enumerable.Range(0, 34).Select(i => i == 10 ? (decimal?)15 : null).ToArray(),
        ClinicBillAmount = "45.00", Ffs = "N", ProcessMonth = "08/2021", Claimability = "A - Accepted", Medications = [new("PANADOL", "3", "10", "5.00")],
    };

    [Fact]
    public async Task Go_add_verify_and_a_rolled_back_save()
    {
        if (TestDb.Cis is null) return;
        await SqlScripts.EnsureDeployedAsync(TestDb.Cis);
        var service = Service();
        Assert.NotEmpty((await service.SetupAsync()).ReceiptModes);
        Assert.Equal(("KLINIK TAN (TAMAN SETIA INDAH)", "G"), ((await service.ClinicAsync("t301")).Name, (await service.ClinicAsync("T301")).ClinicType));
        await Assert.ThrowsAsync<NotFoundException>(() => service.ClinicAsync("ZZZZ"));
        await Assert.ThrowsAsync<NotFoundException>(() => service.DiagnosisAsync("ZZZZ"));

        var member = await service.MemberAsync("DCHC9086608", "00");
        Assert.Equal(("NUR RABIZAH BINTI ADENI", "P"), (member.Member.PatientName, member.MemberType));
        Assert.False(string.IsNullOrEmpty(member.AnnualLimit));

        var add = await service.AddAsync(new ClinicInvoiceRequest(Entry(), true), "devadmin");
        var line = Assert.IsType<ClinicInvoiceLine>(add.Line);
        Assert.Equal((45m, 30m, 15m), (line.Total, line.Amounts[1], line.Amounts[10]));
        Assert.Equal(45m, line.Worksheet.TotalActual);
        var errors = await service.VerifyAsync(new ClinicInvoiceRequest(Entry(), true), "devadmin");
        Assert.Contains("45 Days LATE ! ", errors);

        var repo = new ClinicInvoiceRepository();
        await using var s = await TestDb.Sessions().BeginAsync();
        var caseNo = await repo.NextCaseNoAsync(s);
        Assert.Matches("^EA[0-9]{8}$", caseNo);
        await repo.SaveAsync(s, ClinicInvoiceService.SaveModel(caseNo, Entry(), line, "devadmin", new DateTime(2026, 11, 2)));
        var cas = await s.Sql().Connection.QuerySingleAsync<(string Member, string Clinic, string Invoice)>(
            "SELECT MBMnumber, RTRIM(ClinicCode), RTRIM(InvoiceNo) FROM dbo.CISCAS WHERE CASNumber = @caseNo", new { caseNo }, s.Sql().Transaction);
        Assert.Equal(("DCHC9086608", "T301", "ZZTEST-INV"), cas);
        var claim = await s.Sql().Connection.QuerySingleAsync<decimal>(
            "SELECT CLMTotalActual FROM CISWorksheet.dbo.Claim WHERE CASNumber = @caseNo", new { caseNo }, s.Sql().Transaction);
        Assert.Equal(45m, claim);
        Assert.Equal(1, await s.Sql().Connection.ExecuteScalarAsync<int>(
            "SELECT COUNT(*) FROM CISWorksheet.dbo.ClaimInvoiceActualNew WHERE CASNumber = @caseNo", new { caseNo }, s.Sql().Transaction));
    }
}
