using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Application.Lookups;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IInvoiceAdjustmentRepository"/> on <c>cis.InvAdj_*</c> and the desktop's INVAdj procedures (called with parameters).</summary>
public sealed class InvoiceAdjustmentRepository : IInvoiceAdjustmentRepository
{
    /// <summary>The worksheet procedures and the first worksheet slot each stores.</summary>
    private static readonly (string Proc, int Start)[] Slots =
    [
        ("cis.Ws_INVAdj_SaveINVCLNApp", 426), ("cis.Ws_INVAdj_SaveINVCLNContract", 211), ("cis.Ws_INVAdj_SaveINVActual", 168),
        ("cis.Ws_INVAdj_SaveINVCLNExcess", 383), ("cis.Ws_INVAdj_SaveINVMBMApp", 254), ("cis.Ws_INVAdj_SaveINVMBMExcess", 340),
        ("cis.Ws_INVAdj_SaveINVMBMLmt", 297),
    ];

    private static readonly string[] SlotParams =
    [
        "CLMGPFeeBasicMed", "CLMGPFeeMedInj", "CLMGPFeeLTMed", "CLMSPFeeBasicMed", "CLMSPFeeLTMed", "CLMSPFeeMedInj", "CLMXRay", "CLMUltraSound", "CLMECG",
        "CLMDressing", "CLMNebulizer", "CLMID", "CLMTSSTO", "CLMMinorSurgery", "CLMPOthers", "CLMBloodDrawing", "CLMBloodGlucose", "CLMUrine", "CLMPapSmear",
        "CLMLOthers", "CLMDTPPolio", "CLMDTPPolioHIB", "CLMHepatitisB", "CLMMMR", "CLMIOthers", "CLMDental", "CLMMaternity", "CLMPRVCheckup", "CLMVision",
        "CLMTyphiadJabNoDed", "CLMOthersNoDed",
    ];

    /// <summary>INVAdj_saveActual's items A1..A28 (rows 15–42).</summary>
    private static readonly string[] ActualParams =
    [
        "CLMGPFeeBasicMed", "CLMGPFeeLTMed", "CLMGPFeeMedInj", "CLMSPFeeBasicMed", "CLMSPFeeLTMed", "CLMSPFeeMedInj", "CLMDental", "CLMMaternity",
        "CLMPRVCheckup", "CLMXRay", "CLMUltraSound", "CLMECG", "CLMDressing", "CLMNebulizer", "CLMID", "CLMTSSTO", "CLMMinorSurgery", "CLMPOthers",
        "CLMBloodDrawing", "CLMBloodGlucose", "CLMUrine", "CLMPapSmear", "CLMLOthers", "CLMDTPPolio", "CLMDTPPolioHIB", "CLMHepatitisB", "CLMMMR", "CLMIOthers",
    ];

    private static async Task<List<Dictionary<string, object?>>> Read(SqlMapper.GridReader g) =>
        (await g.ReadAsync()).Select(r => ((IDictionary<string, object?>)r).ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase)).ToList();

    private static string T(Dictionary<string, object?>? r, string k) => InvoiceAdjustmentService.T(r, k);

    /// <inheritdoc />
    public async Task<(Dictionary<string, object?>? Case, Dictionary<string, object?>? Claim, bool Paid, bool Posted, List<string> ReasonIds, decimal? CoPay,
        string Email, string SpecialRemark, bool Btbg)> LoadAsync(IDbSession session, string casNumber)
    {
        using var g = await session.MultipleAsync("cis.InvAdj_Load", new { casNumber });
        var c = (await Read(g)).FirstOrDefault();
        var claim = (await Read(g)).FirstOrDefault();
        var flags = (await Read(g)).FirstOrDefault();
        var ids = (await Read(g)).Select(r => T(r, "ReasonID")).Where(v => v.Length > 0).ToList();
        var copay = (await Read(g)).FirstOrDefault();
        var email = T((await Read(g)).FirstOrDefault(), "MBMEmail");
        var plan = (await Read(g)).FirstOrDefault();
        return (c, claim, T(flags, "Paid") is "1" or "True", T(flags, "Posted") is "1" or "True", ids,
            copay is null ? null : ClinicInvoiceRules.NumOrNull(T(copay, "Copayamt")), email, T(plan, "PLNSpecialRemark"), T(plan, "PLNBTBG") is "1" or "True");
    }

    /// <inheritdoc />
    public async Task<List<LookupOption>> PendingReasonsAsync(IDbSession session, string category) =>
        (await session.RowsAsync("cis.InvAdj_PendingReasons", new { category }))
        .Select(r => new LookupOption(T(r, "PendReasonCode"), T(r, "PendReasonDesc"))).ToList();

    /// <inheritdoc />
    public async Task<List<string>> SameVisitCasesAsync(IDbSession session, string casNumber) =>
        (await session.RowsAsync("cis.InvAdj_SameVisitCases", new { casNumber })).Select(r => T(r, "CasNumber")).ToList();

    /// <inheritdoc />
    public async Task<List<string>> DuplicateInvoicesAsync(IDbSession session, string casNumber, DateTime invoiceDate, string billAmt, string invoiceNo,
        string patientName, string patientIc, string userCode, string coverId)
    {
        var rows = await session.RowsAsync("cis.InvAdj_DuplicateInvoices", new
        {
            casNumber, invoiceDate, billAmt, invoiceNo, patientName, patientIc, userCode, coverId,
        });
        return rows.Count > 0 && T(rows[0], "status") == "1" ? rows.Select(r => T(r, "CasNumber")).ToList() : [];
    }

    /// <inheritdoc />
    public async Task<string> DoubleEntryAsync(IDbSession session, string mbmNumber, string coverId, string clinicCode, string invoiceNo, string visitTime,
        DateTime visit, decimal totalActual, string casNumber) =>
        T((await session.RowsAsync("cis.InvAdj_DoubleEntry", new
        {
            mbmNumber, coverId, clinicCode, invoiceNo, visitTime, dateVisit = visit, totalActual, casNumber,
        })).LastOrDefault(), "isExist");

    /// <inheritdoc />
    public Task<List<InvClaim>> OtherClaimsAsync(IDbSession session, string mbmNumber, int mode, string covId, string casNumber) =>
        session.QueryAsync<InvClaim>("cis.InvAdj_OtherClaims", new { mbmNumber, mode, covId, casNumber });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>?>> WorksheetAsync(IDbSession session, string casNumber) =>
        new InvoiceEnquiryRepository().WorksheetAsync(session, casNumber);

    /// <inheritdoc />
    public Task DiagnosisAsync(IDbSession session, string casNumber, IReadOnlyList<string?> diagnoses, string other, string remark)
    {
        string D(int i) => i < diagnoses.Count ? (diagnoses[i] ?? "").Trim() : "";
        return session.ExecAsync("cis.InvAdj_Diagnosis", new { casNumber, d1 = D(0), d2 = D(1), d3 = D(2), d4 = D(3), d5 = D(4), other, remark });
    }

    /// <inheritdoc />
    public async Task SaveAsync(IDbSession session, AdjustmentSave x)
    {
        var r = x.Request;
        var e = r.Entry;
        var o = r.Extra;
        var a = x.Amounts;
        var cas = r.CaseNo.Trim().ToUpperInvariant();
        var v = x.Worksheet.RowValues;
        decimal R(int i) => v.TryGetValue(i, out var val) && val is decimal d ? d : 0;
        string S(string? t) => (t ?? "").Trim();
        string F(string? t) => S(t).Length > 0 ? S(t)[..1].ToUpperInvariant() : "";
        string Diag(int i) => i < e.Diagnoses.Count ? S(e.Diagnoses[i]) : "";
        InvoiceMedication Med(int i) => i < e.Medications.Count ? e.Medications[i] : new InvoiceMedication(null, null, null, null);
        decimal G(IReadOnlyList<decimal?> l, int i) => i < l.Count ? l[i] ?? 0 : 0;
        var clinic = S(e.ClinicCode).ToUpperInvariant();
        var invoiceDate = ClinicInvoiceRules.Date(e.InvoiceDate);
        var visit = ClinicInvoiceRules.Date(e.VisitDate);
        var mode = S(e.ReceiptMode) is var rm && rm.IndexOf('-') is var dash and > 0 ? rm[..dash].Trim() : rm;
        var time = TimeSpan.TryParse(e.VisitTime, CultureInfo.InvariantCulture, out var ts) ? ts.ToString(@"hh\:mm", CultureInfo.InvariantCulture) : "";
        var payTo = F(o.PayTo) is { Length: > 0 } p ? p : T(x.Loaded.Case, "INVPayTo");
        var bill = ClinicInvoiceRules.Num(e.ClinicBillAmount);
        var m = x.Loaded.Member?.Member;
        var bordxFees = ClinicInvoiceRules.NumOrNull(m?.BordxFees);
        // After posting a new sub-case is written; the loaded case is its master.
        var post = x.NewCase is not null;
        var target = x.NewCase ?? cas;

        var c = new DynamicParameters(new
        {
            CASNumber = target, CASStatus = F(o.Status), Claimability = F(o.Claimability), Pending = F(o.Pending), LastUpdateDate = x.Today, LastUpdateBy = x.UserCode,
            ClinicCode = clinic, NPClinic = clinic is "NP" or "NPS" ? S(o.NpClinic).ToUpperInvariant() : "", InvoiceNo = S(e.InvoiceNo), InvoiceDate = invoiceDate,
            DateVisit = visit, MCDays = e.McDays ?? 0, TimeVisit = time, InvoiceReceiptDate = ClinicInvoiceRules.Date(e.ReceiptDate), InvoiceReceiptMode = mode,
            CasRemark = S(e.Remark), CASTreatReason1 = Diag(0), CASTreatReason2 = Diag(1), CASTreatReason3 = Diag(2), CASTreatReason4 = Diag(3), CASTreatReason5 = Diag(4),
            CASDiagnosis5 = S(e.DiagnosisOther), MCStatus = (e.McDays ?? 0) > 0 ? 1 : 0, NoofDiag = Enumerable.Range(0, 5).Count(i => Diag(i).Length > 0),
            NoOfMedication = Enumerable.Range(0, 7).Count(i => S(Med(i).Type).Length > 0), RepudiationReason = S(o.RepReason), CASMgmtNotes = S(o.MgmtNotes),
            CASRepAction = x.RepAction, CASRepAmt = ClinicInvoiceRules.NumOrNull(o.RepAmount) ?? 0,
            CASInvMonth = invoiceDate?.ToString("MM", CultureInfo.InvariantCulture) ?? "", CASInvYear = invoiceDate?.Year.ToString(CultureInfo.InvariantCulture) ?? "",
            IsPharmacy = F(o.Pharmacy),
        });
        for (var i = 0; i < 7; i++)
        {
            c.Add($"MedInjType{i + 1}", S(Med(i).Type));
            c.Add($"MedInjDays{i + 1}", S(Med(i).Days));
            c.Add($"MedInjQty{i + 1}", S(Med(i).Qty));
        }
        if (post)
        {
            c.Add("MasterCASNumber", cas);
            c.Add("MBMnumber", T(x.Loaded.Case, "MBMnumber"));
            c.Add("CoverID", T(x.Loaded.Case, "CoverID"));
            await session.ExecAsync("dbo.sp_CreateSubCasNum", c);
        }
        else
        {
            // After bordereaux the desktop leaves the pending note out, so the procedure's default (NULL) clears it.
            c.Add("AddPendNote", x.AfterBordx ? null : S(o.AddPendRemark));
            await session.ExecAsync("dbo.INVAdj_SaveCase_2013V3", c);
        }
        if (!x.AfterBordx && !post)
            await session.ExecAsync("cis.InvAdj_SaveExtras", new
            {
                casNumber = cas, qa = string.Join("|", o.Questions), glAmount = bill, c1 = G(o.MedCharges, 0), c2 = G(o.MedCharges, 1), c3 = G(o.MedCharges, 2),
                c4 = G(o.MedCharges, 3), c5 = G(o.MedCharges, 4), c6 = G(o.MedCharges, 5), c7 = G(o.MedCharges, 6),
            });

        if (post)
            await session.ExecAsync("cis.Ws_Save_Invoice", new
            {
                CasNumber = target, INVEnteredDate = x.Today, INVEnteredBy = x.UserCode, INVLastUpdateBy = x.UserCode, INVLastUpdateDate = x.Today, BillNo = S(e.BillNo),
                BillDate = visit, INVAmount = x.Total, INVPayTo = payTo, INVPayee = payTo == "C" ? clinic : "MEMBER", INVBillAmt = bill,
            });
        else
            await session.ExecAsync("cis.Ws_INVAdj_SaveInvoice", new
            {
                CasNumber = cas, INVLastUpdateBy = x.UserCode, INVLastUpdateDate = x.Today, BillNo = S(e.BillNo), BillDate = visit, INVAmount = x.Total, INVPayTo = payTo,
                INVPayee = payTo == "C" ? clinic : "MEMBER", INVBillAmt = bill,
            });
        if (x.AfterBordx) return;
        if (post)
        {
            await CreateSubCaseRestAsync(session, x, target, cas, payTo, R, F, S);
            return;
        }

        var actual = new DynamicParameters(new
        {
            CasNumber = cas, CLMCoPayment = x.ActCoPay, CLMPOthersDesc = S(e.POthDesc), CLMLOthersDesc = S(e.LOthDesc), CLMIOthersDesc = S(e.IOthDesc),
            CLMVision = a[30], CLMTyphiadJabNoDed = a[32], CLMOthersNoDed = a[33], CLMSPFeeBasicMedGSTAmt = 0m, CLMSPFeeLTMedGSTAmt = 0m,
            CLMSPFeeMedInjGSTAmt = 0m, CLMCoPaymentGSTAmt = 0m, CLMVisionGSTAmt = G(o.AGst, 30), CLMTyphiadJabNoDedGSTAmt = G(o.AGst, 32),
            CLMOthersNoDedGSTAmt = G(o.AGst, 33),
        });
        for (var i = 0; i < ActualParams.Length; i++)
        {
            actual.Add(ActualParams[i], a[i + 1]);
            if (i >= 6) actual.Add(ActualParams[i] + "GSTAmt", G(o.AGst, i + 1));
        }
        await session.ExecAsync("cis.Ws_INVAdj_saveActual", actual);

        decimal B(int i) => i < e.B.Count ? e.B[i] ?? 0 : 0;
        await session.ExecAsync("cis.Ws_INVAdj_saveActualOthers", new
        {
            CASNumber = cas, GPFee = B(1), GPLTMed = B(2), GPInj = B(3), GPMed = B(4), SPFee = B(5), SPLTMed = B(6), SPInj = B(7), SPMed = B(8),
            SPFeeGSTAmt = G(o.BGst, 5), SPLTMedGSTAmt = G(o.BGst, 6), SPInjGSTAmt = G(o.BGst, 7), SPMedGSTAmt = G(o.BGst, 8),
        });

        var nonGp = a[31];
        // The desktop reads its non-GP fee before setting it for reimbursements, so Payable2M has no non-GP part.
        decimal payable2C = 0, payable2M = 0, excess = 0;
        if (payTo == "C")
        {
            excess = R(162);
            payable2C = x.ClinicBill - x.McoProcFee;
        }
        else payable2M = bordxFees is > 0 ? 0 : R(162);
        var finalApprove = R(161) != 0 ? R(161) : x.ClinicBill;
        var grpF = m?.GroupMcoStatus == "F";
        decimal bordx = payTo == "M"
            ? bordxFees is null ? (R(161) != 0 ? R(161) + nonGp : payable2M + nonGp) : bordxFees > 0 ? R(161) : R(158) + nonGp
            : grpF ? (R(161) != 0 ? R(161) + nonGp : payable2M + nonGp) : R(158) + nonGp;
        if (bordx == 0) bordx = x.ClinicBill;
        var month = S(e.ProcessMonth);
        await session.ExecAsync("cis.Ws_INVAdj_SaveINVClaim_2015", new
        {
            CasNumber = cas, ClaimCloseDate = x.Today, ClaimClosedBy = x.UserCode, ClaimLastUpdateDate = x.Today, ClaimLastUpdateUser = x.UserCode,
            CLMPayableByMedix2C = payable2C, CLMExcess = excess, CLMMCOProcFee = x.McoProcFee, CLMPayableByMedix2M = payable2M, CLMCLNExcess = R(159),
            CLMTotalActual = bill, CLMFinalApprove = finalApprove, CLMBalanceLimit = R(163), CLMPreviousBalance = R(164), CLMPaidBYMedix2M = R(165),
            CLMPaidBYMedix2C = R(166), CLMAdjAnnLmt = e.AdjustedAnnualLimit ?? 0, ClmNonGPFee = nonGp, NonGPFee = nonGp, CLMCoPayColl = Math.Abs(x.ActCoPay),
            CLMMCOProcStatus = F(e.Ffs), CLMProcMth = month.Length >= 2 ? month[..2] : month,
            CLMProcYR = month.Length > 3 ? month.Substring(3, Math.Min(4, month.Length - 3)) : "", CLMCoPayCLNColl = Math.Abs(a[29]), CLMBordxAmt = bordx,
            CLMBordxMemo = F(o.MemoClaim), CLMRefLetter = F(o.RefLetter), CLMLimitExceeded = e.LimitExceeded ?? 0, CLMTotalExcess = e.Excess ?? 0,
            CLMBTBGAmt = o.BtbgAmount ?? 0, BTBGRemarks = S(o.BtbgRemark),
        });

        foreach (var (proc, start) in Slots)
        {
            var sp = new DynamicParameters();
            sp.Add(proc == "cis.Ws_INVAdj_SaveINVActual" ? "CLMNo" : "CasNumber", cas);
            for (var i = 0; i < SlotParams.Length; i++) sp.Add(SlotParams[i], R(start + i));
            await session.ExecAsync(proc, sp);
        }

        await BalanceAsync(session, x, payTo, R);
    }

    /// <summary>UpdateMBMBalance: the previous approval of the loaded case undone, the new one applied.</summary>
    private static async Task BalanceAsync(IDbSession session, AdjustmentSave x, string payTo, Func<int, decimal> R)
    {
        if (x.Loaded.Member is not { } member) return;
        var m = member.Member;
        var a = x.Amounts;
        var bordxFees = ClinicInvoiceRules.NumOrNull(m.BordxFees);
        var prevAdj = InvoiceAdjustmentService.N(x.Loaded.Claim, "CLMAdjAnnLmt");
        var prevFinal = InvoiceAdjustmentService.N(x.Loaded.Claim, "CLMFinalApprove");
        var approved = payTo == "M" ? (bordxFees is > 0 ? R(160) : R(162)) : R(161);
        var ded = a[32] + a[33] + a[8];
        decimal? balance = member.BalanceLimit == "FULL" ? null
            : payTo == "M" || m.GroupMcoStatus == "F" ? prevFinal - (approved - ded)
            : m.GrossStatus == "N" ? prevFinal - approved
            : m.GrossStatus == "Y" ? 0m : null;
        await session.ExecAsync("cis.InvAdj_Balance", new
        {
            mbmNumber = m.MbmNumber, limitCover = m.CoverId == "00" ? null : m.SuppLimitType is "B" or "D" ? "01" : m.SuppLimitType == "C" ? m.CoverId : null,
            adjDelta = (x.Request.Entry.AdjustedAnnualLimit ?? 0) - prevAdj, accumDelta = approved - prevFinal, balanceDelta = balance,
        });
    }

    /// <summary>frmInvoiceAdjAfterPost's SaveRecord after the case and invoice: the clinic invoice's actual / others / worksheet procedures for the
    /// sub-case, sp_CreateSubCasNumClaim (total actual = the worksheet's) and the LG amount, then the member balance.</summary>
    private static async Task CreateSubCaseRestAsync(IDbSession session, AdjustmentSave x, string target, string master, string payTo, Func<int, decimal> R,
        Func<string?, string> F, Func<string?, string> S)
    {
        var e = x.Request.Entry;
        var o = x.Request.Extra;
        var a = x.Amounts;
        var actual = new DynamicParameters(new
        {
            CasNumber = target, CLMPOthersDesc = S(e.POthDesc), CLMLOthersDesc = S(e.LOthDesc), CLMIOthersDesc = S(e.IOthDesc), CLMVision = a[30],
            CLMTyphiadJabNoDed = a[32], CLMOthersNoDed = a[33],
        });
        for (var i = 0; i < ActualParams.Length; i++) actual.Add(ActualParams[i], a[i + 1]);
        await session.ExecAsync("cis.Ws_Save_ClaimActual", actual);
        decimal B(int i) => i < e.B.Count ? e.B[i] ?? 0 : 0;
        await session.ExecAsync("cis.Ws_SaveActualOthers", new
        {
            CasNumber = target, GPFee = B(1), GPLTMed = B(2), GPInj = B(3), GPMed = B(4), SPFee = B(5), SPLTMed = B(6), SPInj = B(7), SPMed = B(8),
        });

        var m = x.Loaded.Member?.Member;
        var bordxFees = ClinicInvoiceRules.NumOrNull(m?.BordxFees);
        var nonGp = a[31];
        decimal payable2C = 0, payable2M = 0, excess = 0;
        if (payTo == "C")
        {
            excess = R(162);
            payable2C = x.ClinicBill - x.McoProcFee;
        }
        else payable2M = bordxFees is > 0 ? 0 : R(162);
        var finalApprove = R(161) != 0 ? R(161) : x.ClinicBill;
        decimal bordx = payTo == "M"
            ? bordxFees is null ? (R(161) != 0 ? R(161) + nonGp : payable2M + nonGp) : bordxFees > 0 ? R(161) : R(158) + nonGp
            : m?.GroupMcoStatus == "F" ? (R(161) != 0 ? R(161) + nonGp : payable2M + nonGp) : R(158) + nonGp;
        if (bordx == 0) bordx = x.ClinicBill;
        var month = S(e.ProcessMonth);
        await session.ExecAsync("cis.Ws_sp_CreateSubCasNumClaim", new
        {
            CasNumber = target, ClaimCloseDate = x.Today, ClaimClosedBy = x.UserCode, ClaimLastUpdateDate = x.Today, ClaimLastUpdateUser = x.UserCode,
            CLMPayableByMedix2C = payable2C, CLMExcess = excess, CLMMCOProcFee = x.McoProcFee, CLMPayableByMedix2M = payable2M, CLMCLNExcess = R(159),
            CLMTotalActual = R(160), CLMFinalApprove = finalApprove, CLMBalanceLimit = R(163), CLMPreviousBalance = R(164), CLMPaidBYMedix2M = R(165),
            CLMPaidBYMedix2C = R(166), CLMAdjAnnLmt = e.AdjustedAnnualLimit ?? 0, ClmNonGPFee = nonGp, CLMCoPayColl = Math.Abs(x.ActCoPay),
            CLMMCOProcStatus = F(e.Ffs), CLMProcMth = month.Length >= 2 ? month[..2] : month,
            CLMProcYR = month.Length > 3 ? month.Substring(3, Math.Min(4, month.Length - 3)) : "", CLMCoPayCLNColl = Math.Abs(a[29]), CLMBordxAmt = bordx,
            CLMBordxMemo = F(o.MemoClaim), CLMRefLetter = F(o.RefLetter), CLMLimitExceeded = e.LimitExceeded ?? 0, CLMTotalExcess = e.Excess ?? 0,
            MasterCASNumber = master,
        });
        await session.ExecAsync("cis.InvAdj_LgAmount", new { casNumber = target, glAmount = R(160) });
        await ClinicInvoiceRepository.SaveSlotsAsync(session, target, x.Worksheet.RowValues, R(197), R(198));
        await BalanceAsync(session, x, payTo, R);
    }

    /// <inheritdoc />
    public Task BatchAsync(IDbSession session, string casNumber, string? opBatchNo, string userCode, DateTime today) =>
        session.ExecAsync("cis.InvAdj_Batch", new { casNumber, opBatchNo, userCode, today });

    /// <inheritdoc />
    public async Task<string> NextBatchAsync(IDbSession session) => (await session.ScalarAsync<string>("cis.OpBatch_Next", new { opCode = "OC" }))!;

    /// <inheritdoc />
    public async Task<bool> NotifyAsync(IDbSession session, string casNumber, string notes, string pendingReasons, string email) =>
        T((await session.RowsAsync("cis.InvAdj_Notify", new { casNumber, notes, pendingReasons, email })).FirstOrDefault(), "Cnt") == "1";

    /// <inheritdoc />
    public Task AuditAsync(IDbSession session, string casNumber, AuditReviewRequest r, DateTime auditDate) =>
        session.ExecAsync("cis.InvAdj_Audit", new
        {
            casNumber, query = (r.Query ?? "").Trim(), answer = (r.Answer ?? "").Trim(), discount = (r.CostAvoidance ?? "").Trim(), remarks = (r.Remarks ?? "").Trim(),
            auditBy = (r.AuditBy ?? "").Trim(), auditDate,
        });

    /// <inheritdoc />
    public Task ClaimsAppAsync(IDbSession session, string casNumber) => session.ExecAsync("cis.InvAdj_ClaimsApp", new { casNumber });

    /// <inheritdoc />
    public async Task<string> VoidBordxAsync(IDbSession session, string casNumber, string groupCompany, string userCode, DateTime today) =>
        (await session.ScalarAsync<string>("cis.InvAdj_VoidBordx", new { casNumber, groupCompany, userCode, today }))!;
}
