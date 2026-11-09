using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Payment;

/// <summary>Payment To Member search (frmPaymentMemberAuto08 SearchRecord).</summary>
public sealed record PayMbmSearchRequest
{
#pragma warning disable CS1591
    public bool Scan { get; init; }
    public string? CaseFrom { get; init; }
    public string? CaseTo { get; init; }
    public string? MbmNo { get; init; }
    public string? MbmName { get; init; }
    public string? Company { get; init; }
    /// <summary>0 All, 1 Paid, 2 Unpaid, 3 Partial, 4 Paid &amp; Partial, 5 Unpaid &amp; Partial, 6 Paid &amp; Payment Void.</summary>
    public int Payment { get; init; }
    public string? PvFrom { get; init; }
    public string? PvTo { get; init; }
    public bool BlankBatch { get; init; }
    public string? OpFrom { get; init; }
    public string? OpTo { get; init; }
    public bool BlankPv { get; init; }
    public string? PvDateFrom { get; init; }
    public string? PvDateTo { get; init; }
    public bool BlankPvDate { get; init; }
    public string? ChequeFrom { get; init; }
    public string? ChequeTo { get; init; }
    public bool BlankCheque { get; init; }
    public string? CrFrom { get; init; }
    public string? CrTo { get; init; }
    public bool BlankCr { get; init; }
    public string? MonthFrom { get; init; }
    public string? MonthTo { get; init; }
    public bool BlankMonth { get; init; }
    public string? YearFrom { get; init; }
    public string? YearTo { get; init; }
    public bool BlankYear { get; init; }
    /// <summary>0 Not Bordx, 1 Bordx (null: not chosen).</summary>
    public int? Bordx { get; init; } = 1;
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    public int? Btbg { get; init; }
    public string? BtbgFrom { get; init; }
    public string? BtbgTo { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PayMbmSearchRequestValidator : Common.RequestValidator<PayMbmSearchRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayMbmSearchRequestValidator()
    {
        RuleFor(r => r.CaseFrom).MaximumLength(30);
        RuleFor(r => r.CaseTo).MaximumLength(30);
        RuleFor(r => r.MbmNo).MaximumLength(30);
        RuleFor(r => r.MbmName).MaximumLength(200);
        RuleFor(r => r.Company).MaximumLength(200);
    }
}

/// <summary>A ticked case with its Payment2MBM index (blank when unpaid).</summary>
public sealed record PayMbmTick(string CasNumber, string? Index);

/// <summary>Save (one PV for the ticked cases) / Update (the selected PV).</summary>
public sealed record PayMbmSaveRequest
{
#pragma warning disable CS1591
    public string? PvNo { get; init; }
    public string? Index { get; init; }
    public string? CasNumber { get; init; }
    public string? PvDate { get; init; }
    public string? PayTo { get; init; }
    public string? ChequeNo { get; init; }
    public string? ChequeDate { get; init; }
    public string? ChequeAmt { get; init; }
    public string? Wording { get; init; }
    /// <summary>"code - name".</summary>
    public string? Bank { get; init; }
    public string? Remark { get; init; }
    public string? Month { get; init; }
    public string? Year { get; init; }
    /// <summary>Void payment (Cbovoidpayment 1 - VOID).</summary>
    public bool Void { get; init; }
    public List<PayMbmTick> Ticked { get; init; } = [];
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PayMbmSaveRequestValidator : Common.RequestValidator<PayMbmSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayMbmSaveRequestValidator()
    {
        RuleFor(r => r.ChequeNo).MaximumLength(50);
        RuleFor(r => r.PayTo).MaximumLength(60);
        RuleFor(r => r.Remark).MaximumLength(200);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 5000 && t.All(x => x.CasNumber.Length <= 15));
    }
}

/// <summary>Preview PV (Cases, PayTo, PvNo on the screen) / member PV range (PvFrom, PvTo).</summary>
public sealed record PayMbmVoucherRequest(List<string>? Cases, string? PvFrom, string? PvTo, string? PayTo, string? PvNo);

/// <summary>Lengths.</summary>
public sealed class PayMbmVoucherRequestValidator : Common.RequestValidator<PayMbmVoucherRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayMbmVoucherRequestValidator()
    {
        RuleFor(r => r.Cases).Must(c => c is null || (c.Count <= 5000 && c.All(x => x.Length <= 30)));
        RuleFor(r => r.PvFrom).MaximumLength(30);
        RuleFor(r => r.PvTo).MaximumLength(30);
        RuleFor(r => r.PayTo).MaximumLength(200);
        RuleFor(r => r.PvNo).MaximumLength(30);
    }
}

/// <summary>A member payment row (Savecispayment2mbm).</summary>
public record PayMbmRow(string? Index, string CasNumber, string PvNo, DateTime PvDate, DateTime ChequeDate, string ChequeNo, string PayTo, string Bank, string Account,
    string? Remark, decimal PaidAmt, string User, string Month, string Year, string Status);

/// <summary>Data access.</summary>
public interface IPayMbmRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayMbmSearchRequest r, DateTime? pvDateFrom, DateTime? pvDateTo, string? pvNo, string? caseNo);
    Task<List<Dictionary<string, object?>>> ExtraAsync(IDbSession session, string cases);
    Task SaveAsync(IDbSession session, PayMbmRow row);
    Task PvChequeAsync(IDbSession session, string pvNo, DateTime chequeDate, string chequeNo, decimal amount, string payTo, string one, string two);
    Task<List<Dictionary<string, object?>>> PvRowsAsync(IDbSession session, string pvNo, string? casNumber);
    Task EditAsync(IDbSession session, string index, string casNumber, DateTime pvDate, DateTime chequeDate, string chequeNo, decimal amount, string payTo, string? remark,
        string one, string two);
    Task<List<Dictionary<string, object?>>> PvPrintAsync(IDbSession session, string? cases, string? pvFrom, string? pvTo);
    Task<List<Dictionary<string, object?>>> ChequesAsync(IDbSession session, string from, string to, bool account);
    Task ChequesPrintedAsync(IDbSession session, string from, string to);
#pragma warning restore CS1591
}

/// <summary>Payment > To Member (frmPaymentMemberAuto08 with frmPrintPVRange / frmPrintChequeRange in member mode).</summary>
public interface IPayMbmService
{
    /// <summary>cmdSearch_Click (or the PV just saved / the case just updated).</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(PayMbmSearchRequest request, string? pvNo = null, string? caseNo = null);

    /// <summary>CmdSave_Click (after the confirmation): the PV no.</summary>
    Task<string> SaveAsync(PayMbmSaveRequest request, string user);

    /// <summary>CmdUpdate_Click (after the confirmation): the PV no. or the case relisted.</summary>
    Task<(string? PvNo, string? CaseNo)> UpdateAsync(PayMbmSaveRequest request);

    /// <summary>CmdPreviewPV (ticked cases) / the PV range print (frmPrintPVRange member mode).</summary>
    Task<FileOutput> VoucherAsync(IReadOnlyList<string>? cases, string? pvFrom, string? pvTo, string? payTo, string? pvNo, string user);

    /// <summary>frmPrintPVRange PVMBMListing.</summary>
    Task<List<Dictionary<string, object?>>> PvRangeAsync(PayClnInvRangeRequest request);

    /// <summary>frmPrintChequeRange CmdSearch_Click (member).</summary>
    Task<List<Dictionary<string, object?>>> ChequesAsync(PayClnInvRangeRequest request, bool account);

    /// <summary>cmdPrintChq / ChequePrintingMBM / ChequePrintingMBMAcc.</summary>
    Task<FileOutput> PrintChequesAsync(PayClnInvRangeRequest request, bool account);
}

/// <inheritdoc />
public sealed class PayMbmService(IDbSessions sessions, IPayMbmRepository repository, IPayClinicRepository banks, IReportExporter exporter) : IPayMbmService
{
    /// <summary>Desktop texts.</summary>
    public const string NoWorksheet = "Please Enter Worksheet No", NoRecord = "No record found!", NoSelection = "Please select record", NoCheque = "Please Enter Cheque Number!",
        NoBank = "Please Enter Bank !", NoPvCode = "Please keyin Bank - PV Code ! ", Saved = "Record Saved...", Updated = "Record Updated ! ", Paid = "Payment has been made!",
        Broker = "Please check Payment towards borker", NoRecordUpdate = "Please select the record", NoPayTo = "Please Enter Pay to! ", NoChequeDate = "Please Enter Cheque Date!",
        NoChequeAmt = "Please Enter Cheque Amt!", NoMonth = "Please keyin Account Period Month ! ", NoYear = "Please keyin Account Period Year ! ",
        NoWording = "Please Enter Wording Amount! ", UpdateFailed = "Record Update Failed ! ", NoPvRange = "Please Enter PV No", NoPvFound = "No record found";

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(PayMbmSearchRequest r, string? pvNo = null, string? caseNo = null)
    {
        if (pvNo is null && caseNo is null && r.Scan && string.IsNullOrWhiteSpace(r.CaseFrom)) throw new BusinessRuleException(NoWorksheet);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, r, ClinicInvoiceRules.Date(r.PvDateFrom), ClinicInvoiceRules.Date(r.PvDateTo), pvNo, caseNo);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var extra = (await repository.ExtraAsync(s, string.Join(',', rows.Select(x => T(x, "CASNumber")).Distinct())))
            .ToDictionary(x => T(x, "CASNumber"), StringComparer.OrdinalIgnoreCase);
        foreach (var x in rows)
        {
            var e = extra.GetValueOrDefault(T(x, "CASNumber"));
            x["PatientName"] = e?.GetValueOrDefault("PatientName");
            Shape(x, e is null ? 0 : N(e, "PaidTotal"), e is null ? "" : T(e, "FirstStatus"));
        }
        return rows;
    }

    /// <summary>Payment2MBMListing / ProcessCheckedItem: status, paid, balance, amount to pay and the tick message.</summary>
    public static void Shape(Dictionary<string, object?> x, decimal caseTotal, string firstStatus)
    {
        var status = T(x, "PVPayMBMStatus");
        x["Status"] = status switch { "F" => "Full", "P" => "Partial", _ => "" };
        var payable = N(x, "CLMPayableByMedix2M");
        var paid = status == "P" ? caseTotal : N(x, "MBMPaidAmt");
        x["Paid"] = paid;
        x["Balance"] = status switch { "P" => payable - caseTotal, "" => payable - paid, _ => null };
        // On tick: a case with a payment row takes its balance; without one the payable less the postage fee.
        x["AmountToPay"] = firstStatus.Length > 0 ? (status == "P" ? payable - caseTotal : null) : payable - paid - N(x, "CLMPostageFee");
        x["TickError"] = firstStatus.StartsWith('F') ? Paid : null;
        x["TickNotice"] = T(x, "BusinessChanel").StartsWith('B') ? Broker : null;
    }

    /// <summary>The member PV wording: the amount in words split into two lines of up to 50 characters by words.</summary>
    public static (string One, string Two) Lines(decimal amount)
    {
        var words = AmountInWords.Wording(amount).Trim();
        if (words.StartsWith("And", StringComparison.Ordinal)) words = words.Length > 4 ? words[4..].Trim() : "";
        string one = "", two = "";
        foreach (var w in words.Split(' '))
        {
            if (one.Length + two.Length + w.Length <= 50) one = one.Trim().Length == 0 ? w : (one + " " + w).Trim();
            else two = two.Trim().Length == 0 ? w : (two + " " + w).Trim();
        }
        return (one, two);
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(PayMbmSaveRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(NoSelection);
        if (string.IsNullOrWhiteSpace(r.ChequeNo) && !r.Void) throw new BusinessRuleException(NoCheque);
        var bank = Code(r.Bank);
        if (bank.Length == 0) throw new BusinessRuleException(NoBank);
        await using var tx = await sessions.BeginAsync();
        var pvCode = (await banks.LookupsAsync(tx)).Banks.FirstOrDefault(b => b.Code.Equals(bank, StringComparison.OrdinalIgnoreCase)).PvCode ?? "";
        if (pvCode.Length == 0) throw new BusinessRuleException(NoPvCode);
        // The ticked rows as listed now (amounts and statuses are not taken from the browser).
        var listed = new List<Dictionary<string, object?>>();
        foreach (var t in r.Ticked)
        {
            var rows = await repository.SearchAsync(tx, new PayMbmSearchRequest { Scan = true, CaseFrom = t.CasNumber.Trim(), Bordx = null }, null, null, null, null);
            var row = rows.FirstOrDefault(x => T(x, "PVMBMIndex") == (t.Index ?? "").Trim()) ?? throw new BusinessRuleException(NoSelection);
            var e = (await repository.ExtraAsync(tx, t.CasNumber.Trim())).FirstOrDefault();
            Shape(row, e is null ? 0 : N(e, "PaidTotal"), e is null ? "" : T(e, "FirstStatus"));
            if (row["TickError"] is string err) throw new BusinessRuleException(err);
            listed.Add(row);
        }
        var pvNo = r.Void ? "VOID" : await banks.NextPvAsync(tx, bank, pvCode);
        var today = DateTime.Today;
        var cheque = r.Void ? "VOID" : r.ChequeNo!.Trim();
        decimal amount = 0;
        var payTo = "";
        foreach (var x in listed)
        {
            amount += N(x, "CLMPayableByMedix2M");
            payTo = T(x, "MBMName");
            await repository.SaveAsync(tx, new PayMbmRow(Blank(T(x, "PVMBMIndex")), T(x, "CASNumber"), pvNo, today, today, cheque, payTo, bank, T(x, "MBMBankACNo"),
                Blank(r.Remark), x["AmountToPay"] is decimal a ? a : 0, user, (r.Month ?? "").Trim(), (r.Year ?? "").Trim(), "F"));
        }
        if (!r.Void)
        {
            var (one, two) = Lines(amount);
            await repository.PvChequeAsync(tx, pvNo, today, cheque, amount, payTo, one, two);
        }
        await tx.CommitAsync();
        return pvNo;
    }

    /// <inheritdoc />
    public async Task<(string? PvNo, string? CaseNo)> UpdateAsync(PayMbmSaveRequest r)
    {
        if (r.Ticked.Count > 0) return (null, null);
        if (string.IsNullOrWhiteSpace(r.Index)) throw new BusinessRuleException(NoRecordUpdate);
        if (string.IsNullOrWhiteSpace(r.PayTo)) throw new BusinessRuleException(NoPayTo);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(NoCheque);
        var chequeDate = ClinicInvoiceRules.Date(r.ChequeDate) ?? throw new BusinessRuleException(NoChequeDate);
        if (string.IsNullOrWhiteSpace(r.ChequeAmt)) throw new BusinessRuleException(NoChequeAmt);
        if (Code(r.Bank).Length == 0) throw new BusinessRuleException(NoBank);
        if (string.IsNullOrWhiteSpace(r.Month)) throw new BusinessRuleException(NoMonth);
        if (string.IsNullOrWhiteSpace(r.Year)) throw new BusinessRuleException(NoYear);
        if (string.IsNullOrWhiteSpace(r.Wording)) throw new BusinessRuleException(NoWording);
        var amount = Amount(r.ChequeAmt);
        var (one, two) = Lines(amount);
        var pvDate = ClinicInvoiceRules.Date(r.PvDate) ?? chequeDate;
        await using var tx = await sessions.BeginAsync();
        var rows = await repository.PvRowsAsync(tx, (r.PvNo ?? "").Trim(), Blank(r.CasNumber));
        foreach (var x in rows)
        {
            if (T(x, "CASNumber").Length == 0 || T(x, "PVMBMIndex").Length == 0) throw new BusinessRuleException(UpdateFailed);
            await repository.EditAsync(tx, T(x, "PVMBMIndex"), T(x, "CASNumber"), pvDate, chequeDate, r.ChequeNo.Trim(), amount, r.PayTo.Trim(), Blank(r.Remark), one, two);
        }
        await tx.CommitAsync();
        var isVoid = rows.Count == 0 || rows.Any(x => x.GetValueOrDefault("Void") is true);
        return isVoid ? (null, Blank(r.CasNumber)) : ((r.PvNo ?? "").Trim(), null);
    }

    /// <inheritdoc />
    public async Task<FileOutput> VoucherAsync(IReadOnlyList<string>? cases, string? pvFrom, string? pvTo, string? payTo, string? pvNo, string user)
    {
        var preview = cases is { Count: > 0 };
        if (!preview && (string.IsNullOrWhiteSpace(pvFrom) || string.IsNullOrWhiteSpace(pvTo))) throw new BusinessRuleException(preview ? NoSelection : NoPvRange);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.PvPrintAsync(s, preview ? string.Join(',', cases!) : null, preview ? null : pvFrom!.Trim(), preview ? null : pvTo!.Trim());
        if (rows.Count == 0) throw new NotFoundException(NoPvFound);
        var sheets = (preview ? [rows] : rows.GroupBy(x => T(x, "PVNo")).Select(g => g.ToList()).ToList())
            .Select(g => Sheet(g, preview ? (payTo ?? "").Trim() : T(g[0], "PVPayto"), preview ? (pvNo ?? "").Trim() : T(g[0], "PVNo"), user)).ToList();
        return exporter.Export(sheets[0] with { ExtraSheets = sheets.Skip(1).ToList(), FileName = preview ? "PV_Preview" : $"PV_{pvFrom!.Trim()}_{pvTo!.Trim()}" },
            ReportFileType.Excel);
    }

    /// <summary>PreviewPV / PVPrintingMBM: pay to, member address, PV no., cheque no., PV date, prepared by; per case close date, bordx, payee, payable, postage.</summary>
    private static ReportDocument Sheet(List<Dictionary<string, object?>> rows, string payTo, string pvNo, string user)
    {
        var h = rows[0];
        var lines = rows.Select((x, i) => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
        {
            ["No"] = i + 1, ["CASNumber"] = T(x, "CASNumber"), ["CloseDate"] = x.GetValueOrDefault("ClaimCloseDate"), ["BordxRefNo"] = T(x, "BordxRefNo"),
            ["PayTo"] = T(x, "PVPayto"), ["Payable"] = N(x, "CLMPayableByMedix2M"), ["Postage"] = N(x, "CLMPostageFee"),
            ["Net"] = N(x, "CLMPayableByMedix2M") - N(x, "CLMPostageFee"),
        }).ToList();
        return new ReportDocument("PAYMENT VOUCHER",
        [
            new("No", "No"), new("CASNumber", "Case No"), new("CloseDate", "Close Date", ReportFormat.Date), new("BordxRefNo", "Bordx No"), new("PayTo", "Pay To"),
            new("Payable", "Payable", ReportFormat.Money, true), new("Postage", "Postage", ReportFormat.Money, true), new("Net", "Amount", ReportFormat.Money, true),
        ], lines)
        {
            Subtitles =
            [
                payTo, T(h, "MBMAddress1"), $"{T(h, "MBMAddress2")} {T(h, "MBMAddress3")}".Trim(), $"{T(h, "MBMPostCode")} {T(h, "CTYCode")} {T(h, "STACode")}".Trim(),
                $"PV No: {pvNo}   Cheque No: {T(h, "PVChequeNo")}   PV Date: {(h.GetValueOrDefault("PVDate") as DateTime?)?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}",
                $"Prepared by: {user}",
            ],
        };
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> PvRangeAsync(PayClnInvRangeRequest r)
    {
        if (string.IsNullOrWhiteSpace(r.From) || string.IsNullOrWhiteSpace(r.To)) throw new BusinessRuleException(NoPvRange);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.PvPrintAsync(s, null, r.From.Trim(), r.To.Trim());
        if (rows.Count == 0) throw new NotFoundException(NoPvFound);
        return rows;
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> ChequesAsync(PayClnInvRangeRequest r, bool account)
    {
        if (string.IsNullOrWhiteSpace(r.From) || string.IsNullOrWhiteSpace(r.To)) throw new BusinessRuleException(PayClnInvService.NoChequeRange);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.ChequesAsync(s, r.From.Trim(), r.To.Trim(), account);
        if (rows.Count == 0) throw new NotFoundException(NoPvFound);
        return rows;
    }

    /// <inheritdoc />
    public async Task<FileOutput> PrintChequesAsync(PayClnInvRangeRequest r, bool account)
    {
        var rows = await ChequesAsync(r, account);
        var sheets = rows.Select(x =>
        {
            var date = x.GetValueOrDefault("PVChequeDate") as DateTime?;
            var ibm = T(x, "MBMGroupCompany") == "IBM MALAYSIA SDN BHD";
            IReadOnlyList<IReadOnlyDictionary<string, object?>> fields = account
                ? [Field("Account No", T(x, "PVPayeeACNo")), Field("Payee", T(x, "PVPayto"))]
                : [
                    .. ibm ? [Field("Payee", T(x, "PVPayto")), Field("Account No", T(x, "PVPayeeACNo"))] : new[] { Field("Payee", T(x, "PVPayto")) },
                    Field("Amount in words", T(x, "PVWordingOne")), Field("", T(x, "PVWordingTwo")),
                    Field("Amount", N(x, "PVCLNChequeAmt").ToString("#,##0.00", CultureInfo.InvariantCulture)),
                    Field("Date (DDMMYY)", date?.ToString("ddMMyy", CultureInfo.InvariantCulture) ?? ""),
                ];
            return new ReportDocument($"CHEQUE {T(x, "PVChequeNo")}", [new("Field", "Field"), new("Value", "Value")], fields);
        }).ToList();
        if (!account)
        {
            await using var tx = await sessions.BeginAsync();
            await repository.ChequesPrintedAsync(tx, r.From!.Trim(), r.To!.Trim());
            await tx.CommitAsync();
        }
        return exporter.Export(sheets[0] with { ExtraSheets = sheets.Skip(1).ToList(), FileName = $"Cheques_{r.From!.Trim()}_{r.To!.Trim()}" }, ReportFileType.Excel);
    }

    private static IReadOnlyDictionary<string, object?> Field(string f, string v) => new Dictionary<string, object?> { ["Field"] = f, ["Value"] = v };

    private static string Code(string? v)
    {
        var t = (v ?? "").Trim();
        return t.IndexOf('-') is > 0 and var i ? t[..i].Trim() : t;
    }

    private static decimal Amount(string? v) => decimal.TryParse((v ?? "").Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : 0;

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
