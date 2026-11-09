using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Payment;

/// <summary>A Payment2MBM row written by the old To Member screen.</summary>
public record PayMbmOldRow(bool Edit, string? Index, string CasNumber, string PvNo, DateTime PvDate, DateTime ChequeDate, string ChequeNo, decimal ChequeAmt,
    string PayTo, string Bank, string? Account, string? Remark, decimal PaidAmt, string User, string Month, string Year, string Status, string WordingOne, string WordingTwo);

/// <summary>Save / Update of the old To Member screen (the PV frame is entered by the user).</summary>
public sealed record PayMbmOldSaveRequest
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
    public string? Bank { get; init; }
    public string? Account { get; init; }
    public string? Remark { get; init; }
    public string? Month { get; init; }
    public string? Year { get; init; }
    public List<PayMbmTick> Ticked { get; init; } = [];
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PayMbmOldSaveRequestValidator : Common.RequestValidator<PayMbmOldSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public PayMbmOldSaveRequestValidator()
    {
        RuleFor(r => r.PayTo).MaximumLength(60);
        RuleFor(r => r.ChequeNo).MaximumLength(50);
        RuleFor(r => r.Account).MaximumLength(30);
        RuleFor(r => r.Remark).MaximumLength(200);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 5000 && t.All(x => x.CasNumber.Length <= 15));
    }
}

/// <summary>Data access.</summary>
public interface IPayMbmOldRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayMbmSearchRequest r, DateTime? pvDateFrom, DateTime? pvDateTo, string? pvNo);
    Task SaveAsync(IDbSession session, PayMbmOldRow row);
#pragma warning restore CS1591
}

/// <summary>Payment > To Member (old) (frmPaymentMember).</summary>
public interface IPayMbmOldService
{
    /// <summary>CmdSearch_Click (or the PV just saved).</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(PayMbmSearchRequest request, string? pvNo = null);

    /// <summary>CmdSave_Click (after the confirmation): the PV no.</summary>
    Task<string> SaveAsync(PayMbmOldSaveRequest request, string user);

    /// <summary>CmdUpdate_Click (after the confirmation).</summary>
    Task<string> UpdateAsync(PayMbmOldSaveRequest request, string user);

    /// <summary>CmdPrintPV: the voucher of a PV with each case's patient.</summary>
    Task<FileOutput> VoucherAsync(string? pvNo, string? payTo, string user);
}

/// <inheritdoc />
public sealed class PayMbmOldService(IDbSessions sessions, IPayMbmOldRepository repository, IPayMbmRepository member, IPayClinicRepository banks,
    IReportExporter exporter) : IPayMbmOldService
{
    /// <summary>Desktop texts.</summary>
    public const string NoSelection = "Please select the record", NoPayTo = "Please Enter Pay To!", NoPayToUpd = "Please Enter Pay to! ", NoPvDate = "Please Enter PV Date!",
        BadPvDate = "Invalid PV Date!", NoCheque = "Please Enter Cheque Number!", NoChequeDate = "Please Enter Cheque Date!", NoChequeAmt = "Please Enter Cheque Amt!",
        WrongAmount = "Amount to paid is incorrect\nPlease key in remark for reason.", NoWording = "Please Enter Wording Amount!", NoWordingUpd = "Please Enter Wording Amount! ",
        NoBank = "Please Enter Bank !", NoMonth = "Please keyin Account Period Month ! ", NoYear = "Please keyin Account Period Year ! ",
        NoPvCode = "Please keyin Bank - PV Code ! ", Saved = "Record Saved...", Updated = "Record Updated ! ", NoPrintPv = "Please Enter or Generate PV No",
        NoRecord = "No record found!";

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(PayMbmSearchRequest r, string? pvNo = null)
    {
        if (pvNo is null && r.Scan && string.IsNullOrWhiteSpace(r.CaseFrom)) throw new BusinessRuleException(PayMbmService.NoWorksheet);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, r, ClinicInvoiceRules.Date(r.PvDateFrom), ClinicInvoiceRules.Date(r.PvDateTo), pvNo);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        await ShapeAsync(s, rows);
        return rows;
    }

    private async Task ShapeAsync(IDbSession s, List<Dictionary<string, object?>> rows)
    {
        var extra = (await member.ExtraAsync(s, string.Join(',', rows.Select(x => T(x, "CASNumber")).Distinct())))
            .ToDictionary(x => T(x, "CASNumber"), StringComparer.OrdinalIgnoreCase);
        foreach (var x in rows)
        {
            var e = extra.GetValueOrDefault(T(x, "CASNumber"));
            x["PatientName"] = e?.GetValueOrDefault("PatientName");
            PayMbmService.Shape(x, e is null ? 0 : N(e, "PaidTotal"), e is null ? "" : T(e, "FirstStatus"));
            x["TickNotice"] = null;
        }
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(PayMbmOldSaveRequest r, string user)
    {
        if (string.IsNullOrWhiteSpace(r.PayTo)) throw new BusinessRuleException(NoPayTo);
        if (string.IsNullOrWhiteSpace(r.PvDate)) throw new BusinessRuleException(NoPvDate);
        var pvDate = ClinicInvoiceRules.Date(r.PvDate) ?? throw new BusinessRuleException(BadPvDate);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(NoCheque);
        var chequeDate = ClinicInvoiceRules.Date(r.ChequeDate) ?? throw new BusinessRuleException(NoChequeDate);
        if (string.IsNullOrWhiteSpace(r.ChequeAmt)) throw new BusinessRuleException(NoChequeAmt);
        // No ticked case: refused before a PV no. is taken (the desktop took a PV no. and saved nothing).
        if (r.Ticked.Count == 0) throw new BusinessRuleException(NoSelection);
        await using var tx = await sessions.BeginAsync();
        var listed = new List<Dictionary<string, object?>>();
        foreach (var t in r.Ticked)
        {
            var rows = await repository.SearchAsync(tx, new PayMbmSearchRequest { Scan = true, CaseFrom = t.CasNumber.Trim(), Bordx = null }, null, null, null);
            await ShapeAsync(tx, rows);
            var row = rows.FirstOrDefault(x => T(x, "PVMBMIndex") == (t.Index ?? "").Trim()) ?? throw new BusinessRuleException(NoSelection);
            if (row["TickError"] is string err) throw new BusinessRuleException(err);
            listed.Add(row);
        }
        var amount = Amount(r.ChequeAmt);
        if (amount != listed.Sum(x => x["AmountToPay"] is decimal a ? a : 0) && string.IsNullOrWhiteSpace(r.Remark)) throw new BusinessRuleException(WrongAmount);
        if (string.IsNullOrWhiteSpace(r.Wording)) throw new BusinessRuleException(NoWording);
        var bank = Code(r.Bank);
        if (bank.Length == 0) throw new BusinessRuleException(NoBank);
        if (string.IsNullOrWhiteSpace(r.Month)) throw new BusinessRuleException(NoMonth);
        if (string.IsNullOrWhiteSpace(r.Year)) throw new BusinessRuleException(NoYear);
        var pvCode = (await banks.LookupsAsync(tx)).Banks.FirstOrDefault(b => b.Code.Equals(bank, StringComparison.OrdinalIgnoreCase)).PvCode ?? "";
        if (pvCode.Length == 0) throw new BusinessRuleException(NoPvCode);
        var pvNo = await banks.NextPvAsync(tx, bank, pvCode);
        var (one, two) = AmountInWords.Lines(amount);
        foreach (var x in listed)
            await repository.SaveAsync(tx, new PayMbmOldRow(false, Blank(T(x, "PVMBMIndex")), T(x, "CASNumber"), pvNo, pvDate, chequeDate, r.ChequeNo.Trim(), amount,
                r.PayTo.Trim(), bank, Blank(r.Account), Blank(r.Remark), x["AmountToPay"] is decimal a ? a : 0, user, r.Month.Trim(), r.Year.Trim(), "F", one, two));
        await member.PvChequeAsync(tx, pvNo, chequeDate, r.ChequeNo.Trim(), amount, r.PayTo.Trim(), one, two);
        await tx.CommitAsync();
        return pvNo;
    }

    /// <inheritdoc />
    public async Task<string> UpdateAsync(PayMbmOldSaveRequest r, string user)
    {
        if (r.Ticked.Count > 0) return Updated;
        if (string.IsNullOrWhiteSpace(r.Index)) throw new BusinessRuleException(NoSelection);
        if (string.IsNullOrWhiteSpace(r.PayTo)) throw new BusinessRuleException(NoPayToUpd);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(NoCheque);
        var chequeDate = ClinicInvoiceRules.Date(r.ChequeDate) ?? throw new BusinessRuleException(NoChequeDate);
        if (string.IsNullOrWhiteSpace(r.ChequeAmt)) throw new BusinessRuleException(NoChequeAmt);
        var bank = Code(r.Bank);
        if (bank.Length == 0) throw new BusinessRuleException(NoBank);
        if (string.IsNullOrWhiteSpace(r.Month)) throw new BusinessRuleException(NoMonth);
        if (string.IsNullOrWhiteSpace(r.Year)) throw new BusinessRuleException(NoYear);
        if (string.IsNullOrWhiteSpace(r.Wording)) throw new BusinessRuleException(NoWordingUpd);
        var amount = Amount(r.ChequeAmt);
        var (one, two) = AmountInWords.Lines(amount);
        await using var tx = await sessions.BeginAsync();
        var row = (await repository.SearchAsync(tx, new PayMbmSearchRequest { Scan = true, CaseFrom = (r.CasNumber ?? "").Trim(), Bordx = null }, null, null, null))
            .FirstOrDefault(x => T(x, "PVMBMIndex") == r.Index.Trim()) ?? throw new BusinessRuleException(NoSelection);
        // As the desktop, the row keeps its status and the paid amount shown in the list.
        await ShapeAsync(tx, [row]);
        var pvNo = (r.PvNo ?? "").Trim();
        await repository.SaveAsync(tx, new PayMbmOldRow(true, r.Index.Trim(), T(row, "CASNumber"), pvNo, ClinicInvoiceRules.Date(r.PvDate) ?? chequeDate, chequeDate,
            r.ChequeNo.Trim(), amount, r.PayTo.Trim(), bank, Blank(r.Account), Blank(r.Remark), row["Paid"] is decimal p ? p : 0, user, r.Month.Trim(), r.Year.Trim(),
            T(row, "PVPayMBMStatus"), one, two));
        await member.PvChequeAsync(tx, pvNo, chequeDate, r.ChequeNo.Trim(), amount, r.PayTo.Trim(), one, two);
        await tx.CommitAsync();
        return Updated;
    }

    /// <inheritdoc />
    public async Task<FileOutput> VoucherAsync(string? pvNo, string? payTo, string user)
    {
        if (string.IsNullOrWhiteSpace(pvNo)) throw new BusinessRuleException(NoPrintPv);
        await using var s = await sessions.OpenAsync();
        var rows = await member.PvPrintAsync(s, null, pvNo.Trim(), pvNo.Trim());
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var patients = (await member.ExtraAsync(s, string.Join(',', rows.Select(x => T(x, "CASNumber")))))
            .ToDictionary(x => T(x, "CASNumber"), x => T(x, "PatientName"), StringComparer.OrdinalIgnoreCase);
        var h = rows[0];
        var lines = rows.Select((x, i) => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
        {
            ["No"] = i + 1, ["CASNumber"] = T(x, "CASNumber"), ["CloseDate"] = x.GetValueOrDefault("ClaimCloseDate"), ["BordxRefNo"] = T(x, "BordxRefNo"),
            ["Patient"] = patients.GetValueOrDefault(T(x, "CASNumber")), ["Payable"] = N(x, "CLMPayableByMedix2M"), ["Postage"] = N(x, "CLMPostageFee"),
            ["Net"] = N(x, "CLMPayableByMedix2M") - N(x, "CLMPostageFee"),
        }).ToList();
        return exporter.Export(new ReportDocument("PAYMENT VOUCHER",
        [
            new("No", "No"), new("CASNumber", "Case No"), new("CloseDate", "Close Date", ReportFormat.Date), new("BordxRefNo", "Bordx No"), new("Patient", "Patient"),
            new("Payable", "Payable", ReportFormat.Money, true), new("Postage", "Postage", ReportFormat.Money, true), new("Net", "Amount", ReportFormat.Money, true),
        ], lines)
        {
            Subtitles =
            [
                (payTo ?? "").Trim(), T(h, "MBMAddress1"), $"{T(h, "MBMAddress2")} {T(h, "MBMAddress3")}".Trim(), $"{T(h, "MBMPostCode")} {T(h, "CTYCode")} {T(h, "STACode")}".Trim(),
                $"PV No: {pvNo.Trim()}   Cheque No: {T(h, "PVChequeNo")}   PV Date: {(h.GetValueOrDefault("PVDate") as DateTime?)?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}",
                $"Prepared by: {user}",
            ],
            FileName = $"PV_{pvNo.Trim()}",
        }, ReportFileType.Excel);
    }

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
