using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Payment;

/// <summary>FFS Payment search (FrmPaymentFFS SearchRecord).</summary>
public sealed record PayFfsSearchRequest
{
#pragma warning disable CS1591
    public string? Company { get; init; }
    /// <summary>0 All, 1 Paid, 2 Unpaid.</summary>
    public int Payment { get; init; }
    /// <summary>One bordx no. (ChkBordxNo, the default).</summary>
    public bool Scan { get; init; } = true;
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    public bool BlankBordxDate { get; init; }
    public string? BordxDateFrom { get; init; }
    public string? BordxDateTo { get; init; }
    public bool BlankPv { get; init; }
    public string? PvFrom { get; init; }
    public string? PvTo { get; init; }
    public bool BlankPvDate { get; init; }
    public string? PvDateFrom { get; init; }
    public string? PvDateTo { get; init; }
    public bool BlankMonth { get; init; }
    public string? MonthFrom { get; init; }
    public string? MonthTo { get; init; }
    public bool BlankYear { get; init; }
    public string? YearFrom { get; init; }
    public string? YearTo { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PayFfsSearchRequestValidator : Common.RequestValidator<PayFfsSearchRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayFfsSearchRequestValidator()
    {
        RuleFor(r => r.Company).MaximumLength(200);
        RuleFor(r => r.BordxFrom).MaximumLength(20);
        RuleFor(r => r.BordxTo).MaximumLength(20);
        RuleFor(r => r.PvFrom).MaximumLength(30);
        RuleFor(r => r.PvTo).MaximumLength(30);
    }
}

/// <summary>Save (a new PV for the ticked bordereaux) / Update (the selected PV row).</summary>
public sealed record PayFfsSaveRequest
{
#pragma warning disable CS1591
    public string? PvNo { get; init; }
    public string? Index { get; init; }
    public string? BordxNo { get; init; }
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
    public string? Payee { get; init; }
    /// <summary>The ticked bordx numbers.</summary>
    public List<string> Ticked { get; init; } = [];
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PayFfsSaveRequestValidator : Common.RequestValidator<PayFfsSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayFfsSaveRequestValidator()
    {
        RuleFor(r => r.ChequeNo).MaximumLength(30);
        RuleFor(r => r.Account).MaximumLength(50);
        RuleFor(r => r.Remark).MaximumLength(500);
        RuleFor(r => r.Payee).MaximumLength(200);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 5000 && t.All(x => x.Length <= 20));
    }
}

/// <summary>A PaymentFFS row.</summary>
public record PayFfsRow(string? Index, string BordxNo, string PvNo, DateTime PvDate, DateTime? ChequeDate, string? ChequeNo, decimal ChequeAmt, string Bank, string? Account,
    string? Remark, decimal PaidAmt, string Month, string Year, string Payee, string User);

/// <summary>The listing and the total FFS.</summary>
public record PayFfsResult(List<Dictionary<string, object?>> Rows, decimal TotalFfs);

/// <summary>Data access.</summary>
public interface IPayFfsRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayFfsSearchRequest r, string? pvNo, DateTime?[] dates);
    Task SaveAsync(IDbSession session, PayFfsRow row);
#pragma warning restore CS1591
}

/// <summary>Payment > FFS Payment (FrmPaymentFFS).</summary>
public interface IPayFfsService
{
    /// <summary>CmdSearch_Click (or the PV just saved).</summary>
    Task<PayFfsResult> SearchAsync(PayFfsSearchRequest request, string? pvNo = null);

    /// <summary>CmdSave_Click (after the confirmation): the PV no.</summary>
    Task<string> SaveAsync(PayFfsSaveRequest request, string user);

    /// <summary>CmdUpdate_Click (after the confirmation).</summary>
    Task UpdateAsync(PayFfsSaveRequest request, string user);

    /// <summary>CmdPrintPV (pvNo) / CmdPreviewPV (the ticked bordereaux).</summary>
    Task<FileOutput> VoucherAsync(string? pvNo, IReadOnlyList<string>? bordxNos, string user);
}

/// <inheritdoc />
public sealed class PayFfsService(IDbSessions sessions, IPayFfsRepository repository, IPayClinicRepository banks, IReportExporter exporter) : IPayFfsService
{
    /// <summary>Desktop texts.</summary>
    public const string NoBordx = "Please Enter Bordx Ref. No. ", NoRecord = "No Record Found ! ", NoSelection = "Please select the record", BadPvDate = "Invalid PV Date!",
        NoChequeAmt = "Please Enter Cheque Amt!", WrongAmount = "Amount to paid is incorrect\nPlease key in remark for reason.", NoBank = "Please Enter Bank !",
        NoMonth = "Please keyin Account Period Month ! ", NoYear = "Please keyin Account Period Year ! ", NoPvCode = "Please keyin Bank - PV Code ! ",
        NoPayee = "Please keyin Payee Name ! ", Saved = "Record Saved...", Paid = "Payment has been made!", NoPv = "Please Confirm PV No!",
        NoPrintPv = "Please Enter or Generate PV No";

    /// <inheritdoc />
    public async Task<PayFfsResult> SearchAsync(PayFfsSearchRequest r, string? pvNo = null)
    {
        if (pvNo is null && r.Scan && string.IsNullOrWhiteSpace(r.BordxFrom)) throw new BusinessRuleException(NoBordx);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, r, pvNo, [ClinicInvoiceRules.Date(r.BordxDateFrom), ClinicInvoiceRules.Date(r.BordxDateTo),
            ClinicInvoiceRules.Date(r.PvDateFrom), ClinicInvoiceRules.Date(r.PvDateTo)]);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        foreach (var x in rows) Shape(x);
        return new PayFfsResult(rows, rows.Sum(x => N(x, "CLMMCOProcFee")));
    }

    /// <summary>PaymentFFSListing / ProcessCheckedItem: status, balance (FFS − paid) and the tick message.</summary>
    public static void Shape(Dictionary<string, object?> x)
    {
        x["Status"] = T(x, "PVPayStatus") == "F" ? "Full" : "";
        x["Balance"] = N(x, "CLMMCOProcFee") - N(x, "FFSPaidAmt");
        x["TickError"] = T(x, "PVPayStatus") == "F" ? Paid : null;
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(PayFfsSaveRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(NoSelection);
        var pvDate = ClinicInvoiceRules.Date(r.PvDate) ?? throw new BusinessRuleException(BadPvDate);
        if (string.IsNullOrWhiteSpace(r.ChequeAmt)) throw new BusinessRuleException(NoChequeAmt);
        await using var tx = await sessions.BeginAsync();
        var listed = new List<Dictionary<string, object?>>();
        foreach (var b in r.Ticked.Select(x => x.Trim()).Distinct())
        {
            var row = (await repository.SearchAsync(tx, new PayFfsSearchRequest { Scan = true, BordxFrom = b }, null, new DateTime?[4])).FirstOrDefault()
                ?? throw new BusinessRuleException(NoSelection);
            Shape(row);
            if (row["TickError"] is string e) throw new BusinessRuleException(e);
            listed.Add(row);
        }
        var total = listed.Sum(x => (decimal)x["Balance"]!);
        if (Amount(r.ChequeAmt) != total && string.IsNullOrWhiteSpace(r.Remark)) throw new BusinessRuleException(WrongAmount);
        var bank = Code(r.Bank);
        if (bank.Length == 0) throw new BusinessRuleException(NoBank);
        if (string.IsNullOrWhiteSpace(r.Month)) throw new BusinessRuleException(NoMonth);
        if (string.IsNullOrWhiteSpace(r.Year)) throw new BusinessRuleException(NoYear);
        var pvCode = (await banks.LookupsAsync(tx)).Banks.FirstOrDefault(b => b.Code.Equals(bank, StringComparison.OrdinalIgnoreCase)).PvCode ?? "";
        if (pvCode.Length == 0) throw new BusinessRuleException(NoPvCode);
        if (string.IsNullOrWhiteSpace(r.Payee)) throw new BusinessRuleException(NoPayee);
        var pvNo = await banks.NextPvAsync(tx, bank, pvCode);
        foreach (var x in listed)
            await repository.SaveAsync(tx, Row(r, null, T(x, "BordxRefNo"), pvNo, pvDate, bank, (decimal)x["Balance"]!, user));
        await tx.CommitAsync();
        return pvNo;
    }

    /// <inheritdoc />
    public async Task UpdateAsync(PayFfsSaveRequest r, string user)
    {
        if (r.Ticked.Count > 0) return;
        if (string.IsNullOrWhiteSpace(r.Index)) throw new BusinessRuleException(NoSelection);
        if (string.IsNullOrWhiteSpace(r.PvNo)) throw new BusinessRuleException(NoPv);
        var pvDate = ClinicInvoiceRules.Date(r.PvDate) ?? throw new BusinessRuleException(BadPvDate);
        if (string.IsNullOrWhiteSpace(r.ChequeAmt)) throw new BusinessRuleException(NoChequeAmt);
        var bank = Code(r.Bank);
        if (bank.Length == 0) throw new BusinessRuleException(NoBank);
        if (string.IsNullOrWhiteSpace(r.Month)) throw new BusinessRuleException(NoMonth);
        if (string.IsNullOrWhiteSpace(r.Year)) throw new BusinessRuleException(NoYear);
        await using var tx = await sessions.BeginAsync();
        // As the desktop, the paid amount becomes the cheque amount.
        await repository.SaveAsync(tx, Row(r, r.Index.Trim(), (r.BordxNo ?? "").Trim(), r.PvNo.Trim(), pvDate, bank, Amount(r.ChequeAmt), user));
        await tx.CommitAsync();
    }

    /// <inheritdoc />
    public async Task<FileOutput> VoucherAsync(string? pvNo, IReadOnlyList<string>? bordxNos, string user)
    {
        var print = bordxNos is null || bordxNos.Count == 0;
        if (print && string.IsNullOrWhiteSpace(pvNo)) throw new BusinessRuleException(NoPrintPv);
        await using var s = await sessions.OpenAsync();
        var rows = new List<Dictionary<string, object?>>();
        if (print) rows = await repository.SearchAsync(s, new PayFfsSearchRequest { Scan = false }, pvNo!.Trim(), new DateTime?[4]);
        else
            foreach (var b in bordxNos!) rows.AddRange(await repository.SearchAsync(s, new PayFfsSearchRequest { Scan = true, BordxFrom = b.Trim() }, null, new DateTime?[4]));
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var h = rows[0];
        var lines = rows.Select((x, i) => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
        {
            ["No"] = i + 1, ["BordxRefNo"] = T(x, "BordxRefNo"), ["BordxDate"] = x.GetValueOrDefault("BordxDate"),
            ["Amount"] = print ? N(x, "FFSPaidAmt") : N(x, "CLMMCOProcFee") - N(x, "FFSPaidAmt"),
        }).ToList();
        return exporter.Export(new ReportDocument("PAYMENT VOUCHER - FFS",
        [
            new("No", "No"), new("BordxRefNo", "Bordx No"), new("BordxDate", "Bordx Date", ReportFormat.Date),
            new("Amount", print ? "Amount Paid" : "FFS", ReportFormat.Money, true),
        ], lines)
        {
            Subtitles =
            [
                T(h, "GRPCoName"), T(h, "GRPCoAdd1"), T(h, "GRPCoAdd2"), T(h, "GRPCoAdd3"),
                $"PV No: {(print ? pvNo!.Trim() : "")}   Cheque No: {T(h, "PVChqNo")}   Cheque Date: {(h.GetValueOrDefault("PVChqDate") as DateTime?)?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}",
                $"Prepared by: {user}",
            ],
            FileName = $"PV_FFS_{(print ? pvNo!.Trim() : "Preview")}",
        }, ReportFileType.Excel);
    }

    private static PayFfsRow Row(PayFfsSaveRequest r, string? index, string bordx, string pvNo, DateTime pvDate, string bank, decimal paid, string user) =>
        new(index, bordx, pvNo, pvDate, ClinicInvoiceRules.Date(r.ChequeDate), Blank(r.ChequeNo), Amount(r.ChequeAmt), bank, Blank(r.Account), Blank(r.Remark), paid,
            (r.Month ?? "").Trim(), (r.Year ?? "").Trim(), (r.Payee ?? "").Trim(), user);

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
