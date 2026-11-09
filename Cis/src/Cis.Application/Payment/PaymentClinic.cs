using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Payment;

/// <summary>Payment To Clinic search (frmPaymentClinic SearchRecord).</summary>
public sealed record PayClinicSearchRequest
{
#pragma warning disable CS1591
    public string? Company { get; init; }
    /// <summary>"code - name".</summary>
    public string? Clinic { get; init; }
    /// <summary>0 All, 1 Paid, 2 Unpaid (3 Partial, 4 Partial or Paid).</summary>
    public int Payment { get; init; }
    public bool Scan { get; init; }
    public string? CaseFrom { get; init; }
    public string? CaseTo { get; init; }
    public string? InvoiceNo { get; init; }
    public bool BlankInvoiceNo { get; init; }
    public string? InvFrom { get; init; }
    public string? InvTo { get; init; }
    public bool BlankBatch { get; init; }
    public string? OpFrom { get; init; }
    public string? OpTo { get; init; }
    public bool BlankInvDate { get; init; }
    public string? AmtFrom { get; init; }
    public string? AmtTo { get; init; }
    public bool BlankAmt { get; init; }
    public string? PvFrom { get; init; }
    public string? PvTo { get; init; }
    public bool BlankPv { get; init; }
    public string? PvDateFrom { get; init; }
    public string? PvDateTo { get; init; }
    public bool BlankPvDate { get; init; }
    public string? MthFrom { get; init; }
    public string? MthTo { get; init; }
    public bool BlankMth { get; init; }
    public string? YrFrom { get; init; }
    public string? YrTo { get; init; }
    public bool BlankYr { get; init; }
    public string? SunFrom { get; init; }
    public string? SunTo { get; init; }
    /// <summary>0 Health Connect, 1 Sihat Johor.</summary>
    public int Entity { get; init; }
    /// <summary>Upload: invoice numbers listed one by one (cmdUploadDirect / SearchRecordExcel).</summary>
    public List<string>? InvoiceNos { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PayClinicSearchRequestValidator : Common.RequestValidator<PayClinicSearchRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayClinicSearchRequestValidator()
    {
        RuleFor(r => r.Company).MaximumLength(200);
        RuleFor(r => r.Clinic).MaximumLength(300);
        RuleFor(r => r.CaseFrom).MaximumLength(30);
        RuleFor(r => r.CaseTo).MaximumLength(30);
        RuleFor(r => r.InvoiceNo).MaximumLength(50);
        RuleFor(r => r.InvoiceNos).Must(l => l is null || (l.Count <= 5000 && l.All(x => x.Length <= 50)));
    }
}

/// <summary>The PV frame (Add = new PV for the ticked cases, Update = the edited PV).</summary>
public sealed record PayClinicSaveRequest
{
#pragma warning disable CS1591
    /// <summary>Update: the PV being edited; Add: blank (a new PV no. is generated).</summary>
    public string? PvNo { get; init; }
    public string? PvDate { get; init; }
    public string? ChequeNo { get; init; }
    public string? ChequeDate { get; init; }
    public string? ChequeAmt { get; init; }
    /// <summary>"code - name".</summary>
    public string? Bank { get; init; }
    public string? Account { get; init; }
    public string? Remark { get; init; }
    public string? Month { get; init; }
    public string? Year { get; init; }
    /// <summary>The ticked cases with their Payment2CLN index (blank when unpaid); Add only.</summary>
    public List<PayClinicTick> Ticked { get; init; } = [];
    /// <summary>Update: the selected row's Payment2CLN index.</summary>
    public string? Index { get; init; }
#pragma warning restore CS1591
}

/// <summary>A ticked case.</summary>
public sealed record PayClinicTick(string CasNumber, string? Index);

/// <summary>Lengths.</summary>
public sealed class PayClinicSaveRequestValidator : Common.RequestValidator<PayClinicSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayClinicSaveRequestValidator()
    {
        RuleFor(r => r.PvNo).MaximumLength(30);
        RuleFor(r => r.ChequeNo).MaximumLength(30);
        RuleFor(r => r.Account).MaximumLength(50);
        RuleFor(r => r.Remark).MaximumLength(500);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 5000);
    }
}

/// <summary>Print PV (PvNo) / Preview PV (the ticked Cases).</summary>
public sealed record PayClinicVoucherRequest(string? PvNo, List<string>? Cases);

/// <summary>Lengths.</summary>
public sealed class PayClinicVoucherRequestValidator : Common.RequestValidator<PayClinicVoucherRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayClinicVoucherRequestValidator()
    {
        RuleFor(r => r.PvNo).MaximumLength(30);
        RuleFor(r => r.Cases).Must(c => c is null || (c.Count <= 5000 && c.All(x => x.Length <= 30)));
    }
}

/// <summary>A Payment2CLN row to write.</summary>
public record PayClinicRow(string? Index, bool Edit, string CasNumber, string PvNo, DateTime PvDate, DateTime? ChequeDate, string? ChequeNo, decimal ChequeAmt,
    string Bank, string? Account, string? Remark, decimal PaidAmt, string Month, string Year, string Status, string User);

/// <summary>Data access.</summary>
public interface IPayClinicRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayClinicSearchRequest r, string? pvNo, string? clinic, DateTime?[] dates);
    Task<(List<(string Code, string Name, string PvCode)> Banks, List<string> Companies, List<(string Code, string Name)> Clinics)> LookupsAsync(IDbSession session);
    Task<string> NextPvAsync(IDbSession session, string bank, string pvCode);
    Task SaveAsync(IDbSession session, PayClinicRow row);
    Task<List<Dictionary<string, object?>>> PvRowsAsync(IDbSession session, string? pvNo, string? cases);
#pragma warning restore CS1591
}

/// <summary>Payment > To Clinic (frmPaymentClinic).</summary>
public interface IPayClinicService
{
    /// <summary>ComboListing.</summary>
    Task<(List<(string Code, string Name, string PvCode)> Banks, List<string> Companies, List<(string Code, string Name)> Clinics)> LookupsAsync();

    /// <summary>cmdSearch_Click / cmdUploadDirect_Click; the PV list after a save (pvNo).</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(PayClinicSearchRequest request, string? pvNo = null);

    /// <summary>CmdSave_Click (after the confirmation): the new PV no.</summary>
    Task<string> AddAsync(PayClinicSaveRequest request, string user);

    /// <summary>CmdUpdate_Click (after the confirmation).</summary>
    Task UpdateAsync(PayClinicSaveRequest request, string user);

    /// <summary>CmdPrintPV_Click (pvNo) / CmdPreviewPV_Click (ticked cases): the payment voucher.</summary>
    Task<FileOutput> VoucherAsync(string? pvNo, IReadOnlyList<string>? cases, string user);
}

/// <inheritdoc />
public sealed class PayClinicService(IDbSessions sessions, IPayClinicRepository repository, IReportExporter exporter) : IPayClinicService
{
    /// <summary>Desktop texts.</summary>
    public const string NoWorksheet = "Please Enter Worksheet No", NoRecord = "No record found!", NoSelection = "Please select the record", NoPv = "Please Confirm PV No!",
        BadPvDate = "Invalid PV Date!", NoChequeAmt = "Please Enter Cheque Amt!", NoBank = "Please Enter Bank !", NoMonth = "Please keyin Account Period Month ! ",
        NoYear = "Please keyin Account Period Year ! ", NoPvCode = "Please keyin Bank - PV Code ! ", Paid = "Payment has been made!",
        WrongAmount = "Amount to paid is incorrect\nPlease key in remark for reason.", Saved = "Record Saved...", NoPrintPv = "Please Enter or Generate PV No";

    /// <inheritdoc />
    public async Task<(List<(string Code, string Name, string PvCode)> Banks, List<string> Companies, List<(string Code, string Name)> Clinics)> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(PayClinicSearchRequest r, string? pvNo = null)
    {
        if (pvNo is null && r.Scan && string.IsNullOrWhiteSpace(r.CaseFrom) && r.InvoiceNos is null) throw new BusinessRuleException(NoWorksheet);
        await using var s = await sessions.OpenAsync();
        var rows = new List<Dictionary<string, object?>>();
        if (pvNo is null && r.InvoiceNos is { } invoices)
        {
            // SearchRecordExcel: " Where 0=0 and CISCas.InvoiceNo = x" per uploaded invoice no., appended.
            foreach (var inv in invoices.Select(x => x.Trim()).Where(x => x.Length > 0).Distinct())
                rows.AddRange(await repository.SearchAsync(s, new PayClinicSearchRequest { InvoiceNo = inv, Entity = -1 }, "", null, new DateTime?[4]));
        }
        else
        {
            var clinic = (r.Clinic ?? "").Trim();
            if (clinic.IndexOf('-') is > 0 and var i) clinic = clinic[..i].Trim();
            rows = await repository.SearchAsync(s, r, pvNo, clinic.Length > 0 ? clinic : null,
                [ClinicInvoiceRules.Date(r.InvFrom), ClinicInvoiceRules.Date(r.InvTo), ClinicInvoiceRules.Date(r.PvDateFrom), ClinicInvoiceRules.Date(r.PvDateTo)]);
        }
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        foreach (var x in rows) Shape(x);
        return rows;
    }

    /// <summary>Payment2CLNListing: status text, payable (+ non-GP fee), paid and balance as the desktop list.</summary>
    public static void Shape(Dictionary<string, object?> x)
    {
        var status = T(x, "PVPayCLNStatus");
        x["Status"] = status switch { "F" => "Full", "P" => "Partial", _ => "" };
        var payable2C = N(x, "CLMPayableByMedix2C");
        var payable = payable2C + N(x, "CLMNonGPFee");
        var paid = N(x, "CLNPaidAmt");
        x["Payable"] = payable;
        x["Paid"] = paid;
        // Partial: balance against the payable to clinic only (without the non-GP fee), as the desktop; unpaid: payable - paid; full: none.
        x["Balance"] = status switch { "P" => payable2C - paid, "" => payable - paid, _ => null };
    }

    /// <inheritdoc />
    public async Task<string> AddAsync(PayClinicSaveRequest r, string user)
    {
        var pvDate = ClinicInvoiceRules.Date(r.PvDate) ?? throw new BusinessRuleException(BadPvDate);
        if (string.IsNullOrWhiteSpace(r.ChequeAmt)) throw new BusinessRuleException(NoChequeAmt);
        await using var tx = await sessions.BeginAsync();
        var (banks, _, _) = await repository.LookupsAsync(tx);
        var bank = Code(r.Bank);
        var pvCode = banks.FirstOrDefault(b => b.Code.Equals(bank, StringComparison.OrdinalIgnoreCase)).PvCode ?? "";
        // The amounts come from the listed rows (balance on tick), never from the browser.
        var current = new List<(PayClinicTick Tick, decimal Amount)>();
        foreach (var t in r.Ticked)
        {
            var row = (await repository.SearchAsync(tx, new PayClinicSearchRequest { Scan = true, CaseFrom = t.CasNumber, Entity = -1 }, null, null, new DateTime?[4]))
                .FirstOrDefault(x => T(x, "PVCLNIndex") == (t.Index ?? "").Trim()) ?? throw new BusinessRuleException(NoSelection);
            Shape(row);
            if (T(row, "PVPayCLNStatus") == "F") throw new BusinessRuleException(Paid);
            current.Add((t, row["Balance"] is decimal b ? b : 0));
        }
        var total = current.Sum(c => c.Amount);
        if (Amount(r.ChequeAmt) != total && string.IsNullOrWhiteSpace(r.Remark)) throw new BusinessRuleException(WrongAmount);
        if (bank.Length == 0) throw new BusinessRuleException(NoBank);
        if (string.IsNullOrWhiteSpace(r.Month)) throw new BusinessRuleException(NoMonth);
        if (string.IsNullOrWhiteSpace(r.Year)) throw new BusinessRuleException(NoYear);
        if (pvCode.Length == 0) throw new BusinessRuleException(NoPvCode);
        if (current.Count == 0) throw new BusinessRuleException(NoSelection);
        var pvNo = await repository.NextPvAsync(tx, bank, pvCode);
        foreach (var (t, amount) in current)
            await repository.SaveAsync(tx, Row(r, t.Index, false, t.CasNumber, pvNo, pvDate, bank, amount, "F", user));
        await tx.CommitAsync();
        return pvNo;
    }

    /// <inheritdoc />
    public async Task UpdateAsync(PayClinicSaveRequest r, string user)
    {
        if (string.IsNullOrWhiteSpace(r.Index)) throw new BusinessRuleException(NoSelection);
        if (string.IsNullOrWhiteSpace(r.PvNo)) throw new BusinessRuleException(NoPv);
        var pvDate = ClinicInvoiceRules.Date(r.PvDate) ?? throw new BusinessRuleException(BadPvDate);
        if (string.IsNullOrWhiteSpace(r.ChequeAmt)) throw new BusinessRuleException(NoChequeAmt);
        var bank = Code(r.Bank);
        if (bank.Length == 0) throw new BusinessRuleException(NoBank);
        if (string.IsNullOrWhiteSpace(r.Month)) throw new BusinessRuleException(NoMonth);
        if (string.IsNullOrWhiteSpace(r.Year)) throw new BusinessRuleException(NoYear);
        await using var tx = await sessions.BeginAsync();
        // Every Payment2CLN row of the PV gets the frame's PV details, keeping its own case, status and paid amount.
        foreach (var x in await repository.PvRowsAsync(tx, r.PvNo.Trim(), null))
            await repository.SaveAsync(tx, Row(r, T(x, "PVCLNIndex"), true, T(x, "CASNumber"), r.PvNo.Trim(), pvDate, bank, N(x, "CLNPaidAmt"),
                await StatusAsync(tx, T(x, "CASNumber"), T(x, "PVCLNIndex")), user));
        await tx.CommitAsync();
    }

    private async Task<string> StatusAsync(IDbSession s, string cas, string index) =>
        (await repository.SearchAsync(s, new PayClinicSearchRequest { Scan = true, CaseFrom = cas, Entity = -1 }, null, null, new DateTime?[4]))
            .Where(x => T(x, "PVCLNIndex") == index).Select(x => T(x, "PVPayCLNStatus")).FirstOrDefault() ?? "F";

    /// <inheritdoc />
    public async Task<FileOutput> VoucherAsync(string? pvNo, IReadOnlyList<string>? cases, string user)
    {
        var print = !string.IsNullOrWhiteSpace(pvNo) && (cases is null || cases.Count == 0);
        if (!print && (cases is null || cases.Count == 0)) throw new BusinessRuleException(pvNo is null ? NoSelection : NoPrintPv);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.PvRowsAsync(s, print ? pvNo!.Trim() : null, print ? null : string.Join(',', cases!));
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var head = rows[0];
        var hq = !print && T(head, "HqCode").Length > 0;
        string F(string k) => T(head, hq ? "Hq" + k : "CLN" + k);
        var name = hq ? T(head, "HqName") : T(head, "CLNName");
        var code = hq ? T(head, "HqCode") : T(head, "ClinicCode");
        var lines = rows.Select((x, i) => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
        {
            ["No"] = i + 1, ["BillNo"] = T(x, "BillNo"), ["BillDate"] = x.GetValueOrDefault("BillDate"), ["CASNumber"] = T(x, "CASNumber"),
            ["Ref"] = print ? T(x, "ReceiptNo") : T(x, "BordxRefNo"), ["Receipt"] = T(x, "ReceiptNo"), ["Actual"] = print ? N(x, "CLMTotalActual") : null,
            ["Mco"] = print ? N(x, "CLMMCOProcFee") : null, ["Amount"] = print ? N(x, "CLNPaidAmt") : N(x, "CLMPayableByMedix2C"),
        }).ToList();
        var columns = new List<ReportColumn>
        {
            new("No", "No"), new("BillNo", "Bill No"), new("BillDate", "Bill Date", ReportFormat.Date), new("CASNumber", "Case No."),
            new("Ref", print ? "Receipt No." : "Bordx No"),
        };
        if (print) columns.AddRange([new("Actual", "Bill Amount", ReportFormat.Money, true), new("Mco", "MCO Fee", ReportFormat.Money, true)]);
        else columns.Add(new("Receipt", "Receipt No."));
        columns.Add(new("Amount", print ? "Amount Paid" : "Payable", ReportFormat.Money, true));
        return exporter.Export(new ReportDocument("PAYMENT VOUCHER", columns, lines)
        {
            Subtitles =
            [
                $"{name} ({code})", F("Add1"), $"{F("Add2")},{F("Add3")}", $"{F("PostCode")} {F("City")} {F("State")}".Trim(),
                $"PV No: {(print ? pvNo!.Trim() : T(head, "PVCLNNo"))}   Cheque No: {T(head, "PVCLNChequeNo")}   PV Date: {Date(head.GetValueOrDefault("PVCLNDate"))}",
                $"Prepared by: {user}",
            ],
            FileName = $"PV_{(print ? pvNo!.Trim() : "Preview")}",
        }, ReportFileType.Excel);
    }

    private static PayClinicRow Row(PayClinicSaveRequest r, string? index, bool edit, string cas, string pvNo, DateTime pvDate, string bank, decimal paid, string status,
        string user) =>
        new(string.IsNullOrWhiteSpace(index) ? null : index.Trim(), edit, cas, pvNo, pvDate, ClinicInvoiceRules.Date(r.ChequeDate), Blank(r.ChequeNo), Amount(r.ChequeAmt),
            bank, Blank(r.Account), Blank(r.Remark), paid, r.Month!.Trim(), r.Year!.Trim(), status, user);

    private static string Code(string? v)
    {
        var t = (v ?? "").Trim();
        return t.IndexOf('-') is > 0 and var i ? t[..i].Trim() : t;
    }

    private static decimal Amount(string? v) => decimal.TryParse((v ?? "").Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : 0;

    private static string Date(object? v) => v is DateTime d ? d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
