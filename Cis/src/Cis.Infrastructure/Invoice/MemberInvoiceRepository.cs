using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IMemberInvoiceRepository"/> on <c>cis.InvMember_*</c>, the OC batch sequence and the shared worksheet-slot procedures.</summary>
public sealed class MemberInvoiceRepository : IMemberInvoiceRepository
{
    /// <inheritdoc />
    public async Task<bool> IsIbmAsync(IDbSession session, string mbmNumber)
    {
        var rows = await session.RowsAsync("cis.InvMember_Ibm", new { mbmNumber });
        var number = rows.LastOrDefault()?.GetValueOrDefault("Number");
        return number is not null && Convert.ToString(number, CultureInfo.InvariantCulture)!.Trim() is { Length: > 0 } n && n != "0";
    }

    /// <inheritdoc />
    public async Task<string> NextBatchAsync(IDbSession session) => (await session.ScalarAsync<string>("cis.OpBatch_Next", new { opCode = "OC" }))!;

    /// <inheritdoc />
    public async Task SaveAsync(IDbSession session, MemberCaseSave x)
    {
        var e = x.Request.Entry;
        var o = x.Request.Extra;
        var l = x.Line;
        var a = l.Amounts;
        var v = l.Worksheet.RowValues;
        var m = l.Member;
        decimal R(int i) => v.TryGetValue(i, out var val) && val is decimal d ? d : 0;
        string S(string? t) => (t ?? "").Trim();
        string First(string? t) => S(t).Length > 0 ? S(t)[..1].ToUpperInvariant() : "";
        string Diag(int i) => i < e.Diagnoses.Count ? S(e.Diagnoses[i]) : "";
        InvoiceMedication Med(int i) => i < e.Medications.Count ? e.Medications[i] : new InvoiceMedication(null, null, null, null);
        var visit = ClinicInvoiceRules.Date(e.VisitDate);
        var mode = S(e.ReceiptMode) is var rm && rm.IndexOf('-') is var dash and > 0 ? rm[..dash].Trim() : rm;
        var time = TimeSpan.TryParse(e.VisitTime, CultureInfo.InvariantCulture, out var ts) ? ts.ToString(@"hh\:mm", CultureInfo.InvariantCulture) : null;
        var memo = First(o.MemoClaim);

        var c = new DynamicParameters(new
        {
            casNumber = x.CaseNo, mbmNumber = m.Member.MbmNumber, coverId = m.Member.CoverId, userCode = x.UserCode, today = x.Today, clinicCode = l.Clinic.Code,
            npClinic = S(o.NpClinic).ToUpperInvariant(), invoiceNo = S(e.InvoiceNo), invoiceDate = ClinicInvoiceRules.Date(e.InvoiceDate),
            invoiceStmtAmt = ClinicInvoiceRules.Num(e.TotalInvoiceAmount), dateVisit = visit, mcDays = e.McDays ?? 0, visitTime = time, opBatchNo = x.OpBatchNo,
            receiptDate = ClinicInvoiceRules.Date(e.ReceiptDate), receiptMode = mode, remark = S(e.Remark), r1 = Diag(0), r2 = Diag(1), r3 = Diag(2), r4 = Diag(3),
            r5 = Diag(4), diagnosis5 = l.OtherDiagnosis ? S(e.DiagnosisOther).ToUpperInvariant() : "", mcStatus = (e.McDays ?? 0) > 0 ? 1 : 0,
            noOfDiag = Enumerable.Range(0, 5).Count(i => Diag(i).Length > 0), noOfMedication = Enumerable.Range(0, 7).Count(i => S(Med(i).Type).Length > 0),
            clnType = First(o.ClinicType) is { Length: > 0 } ct ? ct : null, repReason = S(o.RepReason), repAction = S(o.RepAction),
            repAmt = S(o.RepAmount), isPharmacy = First(o.Pharmacy),
        });
        for (var i = 0; i < 7; i++)
        {
            c.Add($"t{i + 1}", S(Med(i).Type));
            c.Add($"d{i + 1}", S(Med(i).Days));
            c.Add($"q{i + 1}", S(Med(i).Qty));
        }
        await session.ExecAsync("cis.InvMember_Case", c);

        await session.ExecAsync("cis.InvMember_Invoice", new
        {
            casNumber = x.CaseNo, userCode = x.UserCode, today = x.Today, billNo = S(e.BillNo).ToUpperInvariant(), billDate = visit, amount = l.Total - a[29],
        });

        var actual = new DynamicParameters(new
        {
            casNumber = x.CaseNo, coPayment = a[29], pOthDesc = S(e.POthDesc), lOthDesc = S(e.LOthDesc), iOthDesc = S(e.IOthDesc), vision = a[30], othersNoDed = a[33],
        });
        for (var i = 1; i <= 28; i++) actual.Add($"a{i}", a[i]);
        await session.ExecAsync("cis.InvMember_Actual", actual);

        decimal B(int i) => i < e.B.Count ? e.B[i] ?? 0 : 0;
        await session.ExecAsync("cis.InvMember_ActualOthers", new
        {
            casNumber = x.CaseNo, gpFee = B(1), gpLtMed = B(2), gpInj = B(3), gpMed = B(4), spFee = B(5), spLtMed = B(6), spInj = B(7), spMed = B(8),
        });

        await session.ExecAsync("cis.InvMember_Claim", new
        {
            casNumber = x.CaseNo, userCode = x.UserCode, today = x.Today, payableM = R(162), clnExcess = R(159), totalActual = R(160), finalApprove = R(161),
            balanceLimit = R(163), previousBalance = R(164), paidM = R(165), paidC = R(166), postageFee = 0m, adjAnnLmt = e.AdjustedAnnualLimit ?? 0,
            nonGpFee = a[31], bordxAmt = (R(161) != 0 ? R(161) : R(162)) + a[31], coPay = a[29], memo = memo == "N" ? "N" : "Y", refLetter = First(o.RefLetter),
            deptName = Blank(m.DeptName), costCenter = Blank(m.CostCenter), deptCode = Blank(m.DeptCode), subsidiary = Blank(m.Subsidiary), grpCoCode = m.GroupCompanyCode,
        });

        // The desktop's recordset inserts leave typhoid / others of the actual slots empty.
        await ClinicInvoiceRepository.SaveSlotsAsync(session, x.CaseNo, v, null, null);

        var limitCover = m.Member.CoverId == "00" ? null : m.Member.SuppLimitType is "B" or "D" ? "01" : m.Member.SuppLimitType == "C" ? m.Member.CoverId : null;
        var approved = ClinicInvoiceRules.NumOrNull(m.Member.BordxFees) is > 0 ? R(160) : R(161);
        await session.ExecAsync("cis.InvMember_Balance", new
        {
            mbmNumber = m.Member.MbmNumber, limitCover, adjAnnLmt = e.AdjustedAnnualLimit ?? 0, approved, deduct = approved - (a[8] + a[33]),
            full = m.BalanceLimit == "FULL", health = m.HealthLimitInd == "Y",
        });
    }

    /// <summary>The member invoice keeps blank claim details blank (the clinic invoice writes "NA").</summary>
    private static string Blank(string v) => v == "NA" ? "" : v;
}
