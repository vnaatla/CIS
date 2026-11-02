using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Application.Lookups;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IClinicInvoiceRepository"/>: <c>cis.InvClinic_*</c> and the desktop's save procedures, called with parameters
/// (the desktop pasted the values into SQL text, which broke for amounts of 1,000 or more).</summary>
public sealed class ClinicInvoiceRepository : IClinicInvoiceRepository
{
    /// <summary>The worksheet slot of each SaveINV* procedure (row index of the first of the 31 grid rows).</summary>
    private static readonly (string Proc, int Start)[] Slots =
    [
        ("cis.Ws_SaveINVCLNApp", 426), ("cis.Ws_SaveINVCLNContract", 211), ("cis.Ws_SaveINVActual", 168), ("cis.Ws_SaveINVCLNExcess", 383),
        ("cis.Ws_SaveINVMBMApp", 254), ("cis.Ws_SaveINVMBMExcess", 340), ("cis.Ws_SaveINVMBMLmt", 297),
    ];

    /// <summary>The 31 money parameters of the SaveINV* procedures, in grid order.</summary>
    private static readonly string[] SlotParams =
    [
        "CLMGPFeeBasicMed", "CLMGPFeeMedInj", "CLMGPFeeLTMed", "CLMSPFeeBasicMed", "CLMSPFeeLTMed", "CLMSPFeeMedInj", "CLMXRay", "CLMUltraSound", "CLMECG",
        "CLMDressing", "CLMNebulizer", "CLMID", "CLMTSSTO", "CLMMinorSurgery", "CLMPOthers", "CLMBloodDrawing", "CLMBloodGlucose", "CLMUrine", "CLMPapSmear",
        "CLMLOthers", "CLMDTPPolio", "CLMDTPPolioHIB", "CLMHepatitisB", "CLMMMR", "CLMIOthers", "CLMDental", "CLMMaternity", "CLMPRVCheckup", "CLMVision",
        "CLMTyphiadJabNoDed", "CLMOthersNoDed",
    ];

    /// <summary>Save_ClaimActual's item parameters for A1..A28 (row 15–42).</summary>
    private static readonly string[] ActualParams =
    [
        "CLMGPFeeBasicMed", "CLMGPFeeLTMed", "CLMGPFeeMedInj", "CLMSPFeeBasicMed", "CLMSPFeeLTMed", "CLMSPFeeMedInj", "CLMDental", "CLMMaternity",
        "CLMPRVCheckup", "CLMXRay", "CLMUltraSound", "CLMECG", "CLMDressing", "CLMNebulizer", "CLMID", "CLMTSSTO", "CLMMinorSurgery", "CLMPOthers",
        "CLMBloodDrawing", "CLMBloodGlucose", "CLMUrine", "CLMPapSmear", "CLMLOthers", "CLMDTPPolio", "CLMDTPPolioHIB", "CLMHepatitisB", "CLMMMR", "CLMIOthers",
    ];

    /// <inheritdoc />
    public async Task<ClinicSetup> SetupAsync(IDbSession session)
    {
        using var g = await session.MultipleAsync("cis.InvClinic_Setup");
        var modes = (await g.ReadAsync<LookupOption>()).ToList();
        var months = (await g.ReadAsync<string>()).ToList();
        return new ClinicSetup(modes, months);
    }

    /// <inheritdoc />
    public async Task<ClinicInfo?> ClinicAsync(IDbSession session, string code)
    {
        var r = (await session.RowsAsync("cis.InvClinic_Clinic", new { clinicCode = code })).FirstOrDefault();
        if (r is null) return null;
        string T(string k) => (r.GetValueOrDefault(k) as string)?.Trim() ?? "";
        var panel = T("Panel") is "C" or "Y";
        return new ClinicInfo(T("Code"), panel ? T("Name") : "", T("ClinicType") is { Length: > 0 } t ? t : "G", T("FfsStatus"), panel,
            r.GetValueOrDefault("McoProcFeePercent") is { } p ? Convert.ToDecimal(p, CultureInfo.InvariantCulture) : null);
    }

    /// <inheritdoc />
    public Task<InvoiceDiagnosis?> DiagnosisAsync(IDbSession session, string code) =>
        session.FirstOrDefaultAsync<InvoiceDiagnosis>("cis.InvClinic_Diagnosis", new { code });

    /// <inheritdoc />
    public async Task<InvMemberExtra> MemberExtraAsync(IDbSession session, string mbmNumber, string covId, string planCode)
    {
        using var g = await session.MultipleAsync("cis.InvClinic_MemberExtra", new { mbmNumber, covId, planCode });
        async Task<List<Dictionary<string, object?>>> Read() =>
            (await g.ReadAsync()).Select(r => ((IDictionary<string, object?>)r).ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase)).ToList();
        var visit = (await Read()).FirstOrDefault();
        var plan = (await Read()).FirstOrDefault();
        var gpsp = (await Read()).FirstOrDefault();
        await Read();
        var claim = (await Read()).FirstOrDefault();
        var aso = await Read();
        return new InvMemberExtra(visit, plan, gpsp, claim, aso);
    }

    /// <inheritdoc />
    public Task<List<InvClaim>> ClaimsAsync(IDbSession session, string mbmNumber, int mode, string covId) =>
        session.QueryAsync<InvClaim>("cis.InvClinic_Claims", new { mbmNumber, mode, covId });

    /// <inheritdoc />
    public async Task<(int SameDayCases, List<string> NotPayable)> ChecksAsync(IDbSession session, string mbmNumber, string covId, DateTime visit, string clinic,
        string diagnosisPlan, IReadOnlyList<string?> diagnoses)
    {
        string? D(int i) => i < diagnoses.Count && !string.IsNullOrWhiteSpace(diagnoses[i]) ? diagnoses[i]!.Trim() : null;
        using var g = await session.MultipleAsync("cis.InvClinic_Checks", new
        {
            mbmNumber, covId, visitDate = visit, clinicCode = clinic, card = "", planCode = diagnosisPlan, diag1 = D(0), diag2 = D(1), diag3 = D(2), diag4 = D(3), diag5 = D(4),
        });
        var same = (await g.ReadAsync()).Count();
        await g.ReadAsync();
        var notPayable = (await g.ReadAsync<string>()).ToList();
        return (same, notPayable);
    }

    /// <inheritdoc />
    public async Task<string> NextCaseNoAsync(IDbSession session) => (await session.ScalarAsync<string>("cis.InvClinic_CaseNo"))!;

    /// <inheritdoc />
    public async Task SaveAsync(IDbSession session, ClinicCaseSave x)
    {
        var e = x.Entry;
        var l = x.Line;
        var a = l.Amounts;
        var v = l.Worksheet.RowValues;
        decimal R(int i) => v.TryGetValue(i, out var o) && o is decimal d ? d : 0;
        string S(string? t) => (t ?? "").Trim();
        string? Diag(int i) => i < e.Diagnoses.Count ? S(e.Diagnoses[i]) : "";
        InvoiceMedication Med(int i) => i < e.Medications.Count ? e.Medications[i] : new InvoiceMedication(null, null, null, null);
        var invoiceDate = ClinicInvoiceRules.Date(e.InvoiceDate);
        var visit = ClinicInvoiceRules.Date(e.VisitDate);
        var mode = S(e.ReceiptMode) is var rm && rm.IndexOf('-') is var dash and > 0 ? rm[..dash].Trim() : rm;
        var diagCount = Enumerable.Range(0, 5).Count(i => Diag(i)!.Length > 0);
        var medCount = Enumerable.Range(0, 7).Count(i => S(Med(i).Type).Length > 0);
        var time = TimeSpan.TryParse(e.VisitTime, CultureInfo.InvariantCulture, out var ts) ? ts.ToString(@"hh\:mm", CultureInfo.InvariantCulture) : "";

        var c = new DynamicParameters(new
        {
            CasNumber = x.CaseNo, MBMNumber = x.MbmNumber, CoverID = x.CoverId, CASStatus = "O", Claimability = S(e.Claimability)[..Math.Min(1, S(e.Claimability).Length)],
            CaseType = "C", Pending = "N", RegisterDate = x.Today, REgisterBy = x.UserCode, LastUpdateDate = x.Today, LastUpdateBy = x.UserCode,
            ClinicCode = l.Clinic.Code, InvoiceNo = S(e.InvoiceNo), InvoiceDate = invoiceDate, InvoiceStmtAmt = ClinicInvoiceRules.Num(e.TotalInvoiceAmount),
            DateVisit = visit, MCDays = e.McDays ?? 0, OPBatchNo = "", TimeVisit = time, InvoiceReceiptDate = ClinicInvoiceRules.Date(e.ReceiptDate),
            InvoiceReceiptMode = mode, CasRemark = S(e.Remark), CASTreatReason1 = Diag(0), CASTreatReason2 = Diag(1), CASTreatReason3 = Diag(2), CASTreatReason4 = Diag(3),
            CASTreatReason5 = Diag(4), CASDiagnosis5 = l.OtherDiagnosis ? S(e.DiagnosisOther).ToUpperInvariant() : "", MCStatus = (e.McDays ?? 0) > 0 ? 1 : 0,
            NoOfDiag = diagCount, NoOfMedication = medCount,
            CASInvMonth = invoiceDate?.ToString("MM", CultureInfo.InvariantCulture) ?? "", CASInvYear = invoiceDate?.Year.ToString(CultureInfo.InvariantCulture) ?? "",
        });
        for (var i = 0; i < 7; i++)
        {
            c.Add($"MedInjType{i + 1}", S(Med(i).Type));
            c.Add($"MedInjDays{i + 1}", S(Med(i).Days));
            c.Add($"MedInjQty{i + 1}", S(Med(i).Qty));
        }
        await session.ExecAsync("dbo.Save_Case_2013", c);

        await session.ExecAsync("cis.Ws_Save_Invoice", new
        {
            CasNumber = x.CaseNo, INVEnteredDate = x.Today, INVEnteredBy = x.UserCode, INVLastUpdateBy = x.UserCode, INVLastUpdateDate = x.Today,
            BillNo = S(e.BillNo).ToUpperInvariant(), BillDate = visit, INVAmount = l.Total, INVPayTo = "C", INVPayee = l.Clinic.Code,
            INVBillAmt = ClinicInvoiceRules.Num(e.ClinicBillAmount),
        });

        var actual = new DynamicParameters(new
        {
            CasNumber = x.CaseNo, CLMPOthersDesc = S(e.POthDesc), CLMLOthersDesc = S(e.LOthDesc), CLMIOthersDesc = S(e.IOthDesc), CLMVision = a[30],
            CLMTyphiadJabNoDed = a[32], CLMOthersNoDed = a[33],
        });
        for (var i = 0; i < ActualParams.Length; i++) actual.Add(ActualParams[i], a[i + 1]);
        await session.ExecAsync("cis.Ws_Save_ClaimActual", actual);

        decimal B(int i) => i < e.B.Count ? e.B[i] ?? 0 : 0;
        await session.ExecAsync("cis.Ws_SaveActualOthers", new
        {
            CasNumber = x.CaseNo, GPFee = B(1), GPLTMed = B(2), GPInj = B(3), GPMed = B(4), SPFee = B(5), SPLTMed = B(6), SPInj = B(7), SPMed = B(8),
        });

        var month = S(e.ProcessMonth);
        await session.ExecAsync("cis.Ws_SaveINVClaim_2015", new
        {
            CasNumber = x.CaseNo, ClaimOpenDate = x.Today, ClaimOpenBy = x.UserCode, ClaimCloseDate = x.Today, ClaimClosedBy = x.UserCode,
            ClaimLastUpdateDate = x.Today, ClaimLastUpdateUser = x.UserCode, CLMCoPayColl = l.ActCoPay, CLMPayableByMedix2C = x.Payable, CLMCLNExcess = R(159),
            CLMTotalActual = x.TotalActual, CLMFinalApprove = x.FinalApprove, CLMExcess = R(162), CLMBalanceLimit = R(163), CLMPreviousBalance = R(164),
            CLMPaidBYMedix2M = R(165), CLMPaidBYMedix2C = R(166), CLMAdjAnnLmt = e.AdjustedAnnualLimit ?? 0, ClmNonGPFee = a[31], CLMMCOProcFee = x.McoProcFee,
            CLMMCOProcStatus = S(e.Ffs)[..Math.Min(1, S(e.Ffs).Length)], CLMProcMth = month.Length >= 2 ? month[..2] : month,
            CLMProcYR = month.Length > 3 ? month.Substring(3, Math.Min(4, month.Length - 3)) : "", CLMCoPayCLNColl = a[29], CLMBordxAmt = x.BordxAmt,
            CLMLimitExceeded = e.LimitExceeded ?? 0, CLMTotalExcess = e.Excess ?? 0, CLMCostCenter = l.Member.CostCenter, CLMDeptCode = l.Member.DeptCode,
            CLMSubsidiary = l.Member.Subsidiary, CLMGrpComCode = l.Member.GroupCompanyCode, CLMDeptName = l.Member.DeptName,
        });

        foreach (var (proc, start) in Slots)
        {
            var actualSlots = proc == "cis.Ws_SaveINVActual";
            var p = new DynamicParameters();
            p.Add(actualSlots ? "CLMNo" : "CasNumber", x.CaseNo);
            for (var i = 0; i < SlotParams.Length; i++)
            {
                if (proc == "cis.Ws_SaveINVMBMApp" && SlotParams[i] == "CLMPapSmear") continue;
                // SaveINVActual takes typhoid / others from the item amounts (row 474 / 475).
                p.Add(SlotParams[i], actualSlots && i >= 29 ? (i == 29 ? a[32] : a[33]) : R(start + i));
            }
            await session.ExecAsync(proc, p);
        }

        var charges = new DynamicParameters(new { CASNumber = x.CaseNo });
        for (var i = 0; i < 7; i++)
        {
            charges.Add($"Med{i + 1}", S(Med(i).Type));
            charges.Add($"MedCharges{i + 1}", ClinicInvoiceRules.NumOrNull(Med(i).Rate));
        }
        await session.ExecAsync("cis.Ws_Save_Charges", charges);

        await session.ExecAsync("cis.InvClinic_Balance", new
        {
            mbmNumber = x.MbmNumber, limitCover = x.LimitCover, adjAnnLmt = e.AdjustedAnnualLimit ?? 0, approved = R(161), deduct = x.BalanceDeduction ?? 0,
            full = x.BalanceDeduction is null, mcDays = e.McDays ?? 0,
        });
        await session.ExecAsync("cis.InvClinic_Duplicates", new
        {
            mbmNumber = x.MbmNumber, covId = x.CoverId, visitDate = visit, clinicCode = l.Clinic.Code, userCode = x.UserCode, today = x.Today,
        });
    }

    /// <inheritdoc />
    public Task AsoUtilisationAsync(IDbSession session, string caseNo) =>
        session.ExecAsync("cis.Hdb_sp_ASOClaimsUtilization", new { CASNumber = caseNo, Version = "NULL", CLMType = "OP" });
}
