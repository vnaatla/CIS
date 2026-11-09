using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Payment;

/// <summary>Payment To Clinic - Invoice search (frmPaymentCLNInvoice SearchRecord).</summary>
public sealed record PayClnInvSearchRequest
{
#pragma warning disable CS1591
    /// <summary>"code - name".</summary>
    public string? Clinic { get; init; }
    public string? YearFrom { get; init; }
    public string? YearTo { get; init; }
    public string? MonthFrom { get; init; }
    public string? MonthTo { get; init; }
    /// <summary>0 All, 1 Paid, 2 Unpaid.</summary>
    public int Payment { get; init; }
    public string? PvFrom { get; init; }
    public string? PvTo { get; init; }
    public bool BlankPv { get; init; }
    public string? PvDateFrom { get; init; }
    public string? PvDateTo { get; init; }
    public bool BlankPvDate { get; init; }
    public string? AcMonthFrom { get; init; }
    public string? AcMonthTo { get; init; }
    public bool BlankAcMonth { get; init; }
    public string? AcYearFrom { get; init; }
    public string? AcYearTo { get; init; }
    public bool BlankAcYear { get; init; }
    /// <summary>0 Health Connect, 1 Sihat Johor.</summary>
    public int Entity { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PayClnInvSearchRequestValidator : Common.RequestValidator<PayClnInvSearchRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayClnInvSearchRequestValidator()
    {
        RuleFor(r => r.Clinic).MaximumLength(300);
        RuleFor(r => r.PvFrom).MaximumLength(30);
        RuleFor(r => r.PvTo).MaximumLength(30);
        RuleFor(r => r.YearFrom).MaximumLength(4);
        RuleFor(r => r.YearTo).MaximumLength(4);
    }
}

/// <summary>A ticked invoice-month row (in list order).</summary>
public sealed record PayClnInvTick(string ClinicCode, string Month, string Year);

/// <summary>Save (new PVs for the ticked rows) / Update (the edited PV).</summary>
public sealed record PayClnInvSaveRequest
{
#pragma warning disable CS1591
    public string? PvNo { get; init; }
    public string? PvDate { get; init; }
    public string? ChequeNo { get; init; }
    public string? ChequeDate { get; init; }
    public string? ChequeAmt { get; init; }
    public string? Wording { get; init; }
    public string? PayTo { get; init; }
    /// <summary>"code - name".</summary>
    public string? Bank { get; init; }
    public string? Remark { get; init; }
    public string? Month { get; init; }
    public string? Year { get; init; }
    public int Entity { get; init; }
    public List<PayClnInvTick> Ticked { get; init; } = [];
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PayClnInvSaveRequestValidator : Common.RequestValidator<PayClnInvSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayClnInvSaveRequestValidator()
    {
        RuleFor(r => r.PvNo).MaximumLength(30);
        RuleFor(r => r.ChequeNo).MaximumLength(30);
        RuleFor(r => r.Remark).MaximumLength(500);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 5000 && t.All(x => x.ClinicCode.Length <= 20 && x.Month.Length <= 2 && x.Year.Length <= 4));
    }
}

/// <summary>PV range / cheque range / single PV or cheque.</summary>
public sealed record PayClnInvRangeRequest(string? From, string? To);

/// <summary>Lengths.</summary>
public sealed class PayClnInvRangeRequestValidator : Common.RequestValidator<PayClnInvRangeRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayClnInvRangeRequestValidator()
    {
        RuleFor(r => r.From).MaximumLength(30);
        RuleFor(r => r.To).MaximumLength(30);
    }
}

/// <summary>A Payment2CLN row of this screen.</summary>
public record PayClnInvRow(string? Index, string CasNumber, string PvNo, DateTime PvDate, DateTime? ChequeDate, string ChequeNo, decimal ChequeAmt, string Bank,
    string? Remark, decimal PaidAmt, string Month, string Year, string Status, string WordingOne, string WordingTwo, string User);

/// <summary>Data access.</summary>
public interface IPayClnInvRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayClnInvSearchRequest r, string? clinic, DateTime? pvDateFrom, DateTime? pvDateTo,
        string? pvNo, string? pvLast);
    Task<List<Dictionary<string, object?>>> ClinicsAsync(IDbSession session, string codes);
    Task<List<Dictionary<string, object?>>> InvoicesAsync(IDbSession session, string? clinic, string? year, string? month, string? pvFrom, string? pvTo);
    Task<List<Dictionary<string, object?>>> CasesAsync(IDbSession session, string clinic, string year, string months, int entity);
    Task<List<Dictionary<string, object?>>> PvRowsAsync(IDbSession session, string pvNo);
    Task SaveAsync(IDbSession session, PayClnInvRow row);
    Task<List<Dictionary<string, object?>>> ChequesAsync(IDbSession session, string from, string to);
    Task ChequesPrintedAsync(IDbSession session, string from, string to);
#pragma warning restore CS1591
}

/// <summary>Payment > To Clinic - Invoice (frmPaymentCLNInvoice with frmInvoiceDetPopup, frmPrintPVRange and frmPrintChequeRange).</summary>
public interface IPayClnInvService
{
    /// <summary>CmdSearch_Click (or the PVs just saved).</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(PayClnInvSearchRequest request, string? pvNo = null, string? pvLast = null);

    /// <summary>lstPaymentCLN_DblClick: the invoices of a row (frmInvoiceDetPopup).</summary>
    Task<List<Dictionary<string, object?>>> InvoicesAsync(string clinic, string month, string year);

    /// <summary>CmdSave_Click (after the confirmation): the first and last PV no.</summary>
    Task<(string First, string Last)> SaveAsync(PayClnInvSaveRequest request, string user);

    /// <summary>CmdUpdate_Click (after the confirmation).</summary>
    Task UpdateAsync(PayClnInvSaveRequest request, string user);

    /// <summary>CmdPrintPV (pvNo) / CmdPreviewPV (ticked rows): the payment voucher.</summary>
    Task<FileOutput> VoucherAsync(string? pvNo, IReadOnlyList<PayClnInvTick>? ticked, string user);

    /// <summary>frmPrintPVRange: the PVs of a range.</summary>
    Task<List<Dictionary<string, object?>>> PvRangeAsync(PayClnInvRangeRequest request);

    /// <summary>frmPrintPVRange cmdPrint: one voucher sheet per PV.</summary>
    Task<FileOutput> PrintPvRangeAsync(PayClnInvRangeRequest request, string user);

    /// <summary>frmPrintChequeRange CmdSearch_Click.</summary>
    Task<List<Dictionary<string, object?>>> ChequesAsync(PayClnInvRangeRequest request);

    /// <summary>cmdPrintChq / frmPrintChequeRange cmdPrint: the cheques, marked printed.</summary>
    Task<FileOutput> PrintChequesAsync(PayClnInvRangeRequest request);
}

/// <inheritdoc />
public sealed class PayClnInvService(IDbSessions sessions, IPayClnInvRepository repository, IPayClinicRepository banks, IReportExporter exporter) : IPayClnInvService
{
    /// <summary>Desktop texts.</summary>
    public const string NoRecord = "No record found!", NoSelection = "Please select the record", NoCheque = "Please Enter Cheque Number!",
        NoPayee = "Please Check Payee! ", NoPayeeAddress = "Please Check Payee's Address 1 ! ", Duplicate = "Please Check Duplicate Records! ",
        Paid = "Payment has been made!", Saved = "Record Saved...", NoPayTo = "Please Enter Pay to! ", NoPv = "Please Confirm PV No!", BadPvDate = "Invalid PV Date!",
        NoChequeDate = "Please Enter Cheque Date!", NoChequeAmt = "Please Enter Cheque Amt!", NoWording = "Please Enter Wording Amount!", NoBank = "Please Enter Bank !",
        NoMonth = "Please keyin Account Period Month ! ", NoYear = "Please keyin Account Period Year ! ", NoPrintPv = "Please Enter or Generate PV No",
        NoPvRange = "Please Enter PV No", NoPvFound = "No record found", NoChequeRange = "Please Enter Cheque No", BadChequeRange = "Please ensure cheque range entered correctly",
        NotInSeries = "Cheque range not in series", NoBankPvCode = "Please keyin Bank - PV Code ! ";

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(PayClnInvSearchRequest r, string? pvNo = null, string? pvLast = null)
    {
        var clinic = Code(r.Clinic);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, r, clinic.Length > 0 ? clinic : null, ClinicInvoiceRules.Date(r.PvDateFrom), ClinicInvoiceRules.Date(r.PvDateTo),
            Blank(pvNo), Blank(pvLast));
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var clinics = (await repository.ClinicsAsync(s, string.Join(',', rows.Select(x => T(x, "ClinicCode")).Distinct())))
            .ToDictionary(x => T(x, "CLNCode"), StringComparer.OrdinalIgnoreCase);
        foreach (var x in rows)
        {
            foreach (var (k, v) in clinics.GetValueOrDefault(T(x, "ClinicCode")) ?? []) if (k != "CLNCode") x[k] = v;
            if (r.Entity == 1) x["CLNMCOProcfeePerc"] = "NA";
            Shape(x);
        }
        return rows;
    }

    /// <summary>Payment2CLNListing / ProcessCheckedItem: status, payable, paid, balance and the tick message.</summary>
    public static void Shape(Dictionary<string, object?> x)
    {
        var status = T(x, "PVPayCLNStatus");
        x["Status"] = status switch { "F" => "Full", "P" => "Partial", _ => "" };
        x["Balance"] = N(x, "CLMPayableByMedix2C") - N(x, "CLNPaidAmt");
        x["TickError"] = T(x, "CLNPName").Length == 0 ? NoPayee : T(x, "CLNPAdd1").Length == 0 ? NoPayeeAddress : status.Length > 0 ? Paid : null;
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> InvoicesAsync(string clinic, string month, string year)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.InvoicesAsync(s, clinic.Trim(), year.Trim(), month.Trim(), null, null);
        foreach (var x in rows) Shape(x);
        return rows;
    }

    /// <inheritdoc />
    public async Task<(string First, string Last)> SaveAsync(PayClnInvSaveRequest r, string user)
    {
        var bank = Code(r.Bank);
        if (r.Ticked.Count == 0) throw new BusinessRuleException(NoSelection);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(NoCheque);
        if (r.Ticked.GroupBy(t => (t.ClinicCode.Trim(), t.Month.Trim(), t.Year.Trim())).Any(g => g.Count() > 1)) throw new BusinessRuleException(Duplicate);
        await using var tx = await sessions.BeginAsync();
        var pvCode = (await banks.LookupsAsync(tx)).Banks.FirstOrDefault(b => b.Code.Equals(bank, StringComparison.OrdinalIgnoreCase)).PvCode ?? "";
        if (r.Entity == 0 && pvCode.Length == 0) throw new BusinessRuleException(NoBankPvCode);
        var today = DateTime.Today;
        string? first = null, last = null;
        var cheque = r.ChequeNo.Trim();
        var count = 0;
        // Consecutive ticked rows of the same clinic make one PV; the group's year is the year of its last row (as the desktop).
        var groups = new List<List<PayClnInvTick>>();
        foreach (var t in r.Ticked)
            if (groups.Count > 0 && groups[^1][0].ClinicCode.Trim() == t.ClinicCode.Trim()) groups[^1].Add(t);
            else groups.Add([t]);
        foreach (var g in groups)
        {
            var clinic = g[0].ClinicCode.Trim();
            var cases = await repository.CasesAsync(tx, clinic, g[^1].Year.Trim(), string.Join(',', g.Select(t => t.Month.Trim())), r.Entity);
            var subTotal = cases.Sum(c => N(c, "CLMPayableByMedix2C"));
            if (subTotal == 0 || (g == groups[^1] && subTotal < 0)) continue;
            count++;
            var pvNo = r.Entity == 1 ? await banks.NextPvAsync(tx, "SJ", "CS") : await banks.NextPvAsync(tx, bank, pvCode);
            if (count > 1) cheque = NextCheque(cheque);
            first ??= pvNo;
            last = pvNo;
            var (one, two) = AmountInWords.Lines(subTotal);
            foreach (var c in cases)
                await repository.SaveAsync(tx, new PayClnInvRow(null, T(c, "CASNumber"), pvNo, today, today, cheque, subTotal, bank, Blank(r.Remark),
                    N(c, "CLMPayableByMedix2C"), (r.Month ?? "").Trim(), (r.Year ?? "").Trim(), "F", one, two, user));
        }
        if (first is null) throw new BusinessRuleException(NoSelection);
        await tx.CommitAsync();
        return (first, last!);
    }

    /// <summary>The next cheque no.: the first 3 characters and the number after them plus one, 6 digits (as the desktop).</summary>
    public static string NextCheque(string cheque)
    {
        var code = cheque.Length >= 3 ? cheque[..3] : cheque;
        var rest = cheque.Length > 3 ? cheque.Substring(3, Math.Min(8, cheque.Length - 3)) : "";
        return code + ((long.TryParse(rest, out var n) ? n : 0) + 1).ToString("000000", CultureInfo.InvariantCulture);
    }

    /// <inheritdoc />
    public async Task UpdateAsync(PayClnInvSaveRequest r, string user)
    {
        if (r.Ticked.Count > 0) return;
        if (string.IsNullOrWhiteSpace(r.PayTo)) throw new BusinessRuleException(NoPayTo);
        if (string.IsNullOrWhiteSpace(r.PvNo)) throw new BusinessRuleException(NoPv);
        var pvDate = ClinicInvoiceRules.Date(r.PvDate) ?? throw new BusinessRuleException(BadPvDate);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(NoCheque);
        var chequeDate = ClinicInvoiceRules.Date(r.ChequeDate) ?? throw new BusinessRuleException(NoChequeDate);
        if (string.IsNullOrWhiteSpace(r.ChequeAmt)) throw new BusinessRuleException(NoChequeAmt);
        if (string.IsNullOrWhiteSpace(r.Wording)) throw new BusinessRuleException(NoWording);
        var bank = Code(r.Bank);
        if (bank.Length == 0) throw new BusinessRuleException(NoBank);
        if (string.IsNullOrWhiteSpace(r.Month)) throw new BusinessRuleException(NoMonth);
        if (string.IsNullOrWhiteSpace(r.Year)) throw new BusinessRuleException(NoYear);
        var amount = Amount(r.ChequeAmt);
        var (one, two) = AmountInWords.Lines(amount);
        await using var tx = await sessions.BeginAsync();
        foreach (var x in await repository.PvRowsAsync(tx, r.PvNo.Trim()))
            await repository.SaveAsync(tx, new PayClnInvRow(T(x, "PVCLNIndex"), T(x, "CASNumber"), r.PvNo.Trim(), pvDate, chequeDate, r.ChequeNo.Trim(), amount, bank,
                Blank(r.Remark), N(x, "CLNPaidAmt"), r.Month.Trim(), r.Year.Trim(), "F", one, two, user));
        await tx.CommitAsync();
    }

    /// <inheritdoc />
    public async Task<FileOutput> VoucherAsync(string? pvNo, IReadOnlyList<PayClnInvTick>? ticked, string user)
    {
        var print = ticked is null || ticked.Count == 0;
        if (print && string.IsNullOrWhiteSpace(pvNo)) throw new BusinessRuleException(NoPrintPv);
        await using var s = await sessions.OpenAsync();
        var doc = print ? await PvSheetAsync(s, pvNo!.Trim(), user) : null;
        if (doc is null && !print)
        {
            // Preview: the ticked rows' invoices with their payable to clinic.
            var rows = new List<Dictionary<string, object?>>();
            foreach (var t in ticked!) rows.AddRange(await repository.InvoicesAsync(s, t.ClinicCode.Trim(), t.Year.Trim(), t.Month.Trim(), null, null));
            rows = rows.Where(x => T(x, "PVCLNNo").Length == 0).ToList();
            if (rows.Count == 0) throw new NotFoundException(NoRecord);
            doc = Sheet(rows, "", user, false);
        }
        return exporter.Export(doc ?? throw new NotFoundException(NoRecord), ReportFileType.Excel);
    }

    private async Task<ReportDocument?> PvSheetAsync(IDbSession s, string pvNo, string user)
    {
        var rows = await repository.InvoicesAsync(s, null, null, null, pvNo, pvNo);
        return rows.Count == 0 ? null : Sheet(rows, pvNo, user, true);
    }

    /// <summary>PrintPV: payee, clinic and payee address, PV no., cheque no., PV date, prepared by; per invoice no., date, approved, MCO fee, amount.</summary>
    private static ReportDocument Sheet(List<Dictionary<string, object?>> rows, string pvNo, string user, bool print)
    {
        var h = rows[0];
        var lines = rows.Select((x, i) => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
        {
            ["No"] = i + 1, ["InvoiceNo"] = T(x, "InvoiceNo"), ["InvoiceDate"] = x.GetValueOrDefault("InvoiceDate"), ["Approved"] = N(x, "CLMFinalApprove"),
            ["Mco"] = N(x, "CLMMCOProcFee"), ["Amount"] = print ? N(x, "CLNPaidAmt") : N(x, "CLMPayableByMedix2C"),
        }).ToList();
        return new ReportDocument("PAYMENT VOUCHER",
        [
            new("No", "No"), new("InvoiceNo", "Invoice No"), new("InvoiceDate", "Invoice Date", ReportFormat.Date), new("Approved", "Approved", ReportFormat.Money, true),
            new("Mco", "MCO Fee", ReportFormat.Money, true), new("Amount", print ? "Amount Paid" : "Payable", ReportFormat.Money, true),
        ], lines)
        {
            Subtitles =
            [
                T(h, "CLNPName"), $"{T(h, "CLNName")} ({T(h, "ClinicCode")})", T(h, "CLNPAdd1"), T(h, "CLNPAdd2"), T(h, "CLNPAdd3"),
                $"PV No: {pvNo}   Cheque No: {T(h, "PVCLNChequeNo")}   PV Date: {Date(h.GetValueOrDefault("PVCLNDate"))}", $"Prepared by: {user}",
            ],
            FileName = $"PV_{(pvNo.Length > 0 ? pvNo : "Preview")}",
        };
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> PvRangeAsync(PayClnInvRangeRequest r)
    {
        if (string.IsNullOrWhiteSpace(r.From) || string.IsNullOrWhiteSpace(r.To)) throw new BusinessRuleException(NoPvRange);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.InvoicesAsync(s, null, null, null, r.From.Trim(), r.To.Trim());
        if (rows.Count == 0) throw new NotFoundException(NoPvFound);
        return rows;
    }

    /// <inheritdoc />
    public async Task<FileOutput> PrintPvRangeAsync(PayClnInvRangeRequest r, string user)
    {
        var pvs = (await PvRangeAsync(r)).Select(x => T(x, "PVCLNNo")).Where(x => x.Length > 0).Distinct().OrderBy(x => x, StringComparer.Ordinal).ToList();
        await using var s = await sessions.OpenAsync();
        var sheets = new List<ReportDocument>();
        foreach (var pv in pvs) if (await PvSheetAsync(s, pv, user) is { } d) sheets.Add(d);
        if (sheets.Count == 0) throw new NotFoundException(NoPvFound);
        return exporter.Export(sheets[0] with { ExtraSheets = sheets.Skip(1).ToList(), FileName = $"PV_{r.From!.Trim()}_{r.To!.Trim()}" }, ReportFileType.Excel);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> ChequesAsync(PayClnInvRangeRequest r)
    {
        if (string.IsNullOrWhiteSpace(r.From) || string.IsNullOrWhiteSpace(r.To)) throw new BusinessRuleException(NoChequeRange);
        var selected = Seq(r.To) - Seq(r.From) + 1;
        if (selected <= 0) throw new BusinessRuleException(BadChequeRange);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.ChequesAsync(s, r.From.Trim(), r.To.Trim());
        if (rows.Count == 0) throw new NotFoundException(NoPvFound);
        if (rows.Count != selected) throw new BusinessRuleException(NotInSeries);
        return rows;
    }

    /// <inheritdoc />
    public async Task<FileOutput> PrintChequesAsync(PayClnInvRangeRequest r)
    {
        var rows = await ChequesAsync(r);
        // One sheet per cheque with the cheque's fields: payee, the two wording lines, the amount and the date as DDMMYY digits.
        var sheets = rows.Select(x =>
        {
            var date = x.GetValueOrDefault("PVCLNChequeDate") as DateTime?;
            return new ReportDocument($"CHEQUE {T(x, "PVCLNChequeNo")}", [new("Field", "Field"), new("Value", "Value")],
            [
                Field("Payee", T(x, "CLNPName")), Field("Amount in words", T(x, "PVWordingOne")), Field("", T(x, "PVWordingTwo")),
                Field("Amount", N(x, "PVCLNChequeAmt").ToString("#,##0.00", CultureInfo.InvariantCulture)),
                Field("Date (DDMMYY)", date?.ToString("ddMMyy", CultureInfo.InvariantCulture) ?? ""),
            ]);
        }).ToList();
        await using var tx = await sessions.BeginAsync();
        await repository.ChequesPrintedAsync(tx, r.From!.Trim(), r.To!.Trim());
        await tx.CommitAsync();
        return exporter.Export(sheets[0] with { ExtraSheets = sheets.Skip(1).ToList(), FileName = $"Cheques_{r.From.Trim()}_{r.To.Trim()}" }, ReportFileType.Excel);
    }

    private static IReadOnlyDictionary<string, object?> Field(string f, string v) => new Dictionary<string, object?> { ["Field"] = f, ["Value"] = v };

    private static long Seq(string? cheque)
    {
        var t = (cheque ?? "").Trim();
        return t.Length > 3 && long.TryParse(t.Substring(3, Math.Min(8, t.Length - 3)), out var n) ? n : 0;
    }

    private static string Code(string? v)
    {
        var t = (v ?? "").Trim();
        return t.IndexOf('-') is > 0 and var i ? t[..i].Trim() : t;
    }

    private static decimal Amount(string? v) => decimal.TryParse((v ?? "").Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : 0;

    private static string Date(object? v) => v is DateTime d ? d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : "";

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
