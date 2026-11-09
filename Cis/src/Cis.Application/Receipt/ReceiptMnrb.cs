using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Receipt;

/// <summary>Receipt from MNRB search (FrmReceiptMNRB SearchRecord).</summary>
public sealed record RcpMnrbSearchRequest
{
#pragma warning disable CS1591
    /// <summary>false: by bordereaux (OptBordx), true: by worksheet (OptWks).</summary>
    public bool Worksheet { get; init; }
    public bool Scan { get; init; } = true;
    /// <summary>"code - name" (HISDB PAY, HLTCode S).</summary>
    public string? Payor { get; init; }
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    public string? CaseFrom { get; init; }
    public string? CaseTo { get; init; }
    /// <summary>0 All, 1 Paid, 2 Unpaid.</summary>
    public int WsStatus { get; init; }
    /// <summary>0 All, 1 Paid, 2 Unpaid, 3 Partial, 4 Paid &amp; Partial, 5 Unpaid &amp; Partial.</summary>
    public int BordxStatus { get; init; }
    public string? ReceiptFrom { get; init; }
    public string? ReceiptTo { get; init; }
    public string? ReceiptDateFrom { get; init; }
    public string? ReceiptDateTo { get; init; }
    public string? MonthFrom { get; init; }
    public string? MonthTo { get; init; }
    public string? YearFrom { get; init; }
    public string? YearTo { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class RcpMnrbSearchRequestValidator : Common.RequestValidator<RcpMnrbSearchRequest>
{
    /// <summary>Defines the rules.</summary>
    public RcpMnrbSearchRequestValidator()
    {
        RuleFor(r => r.Payor).MaximumLength(300);
        RuleFor(r => r.BordxFrom).MaximumLength(20);
        RuleFor(r => r.BordxTo).MaximumLength(20);
        RuleFor(r => r.CaseFrom).MaximumLength(30);
        RuleFor(r => r.CaseTo).MaximumLength(30);
    }
}

/// <summary>Save (the ticked rows) / Update (the selected row).</summary>
public sealed record RcpMnrbSaveRequest
{
#pragma warning disable CS1591
    public bool Worksheet { get; init; }
    public string? ReceiptNo { get; init; }
    public string? ReceiptDate { get; init; }
    public string? ChequeNo { get; init; }
    public string? ReceiptAmt { get; init; }
    public string? Remark { get; init; }
    public string? Month { get; init; }
    public string? Year { get; init; }
    public List<RcpPayTick> Ticked { get; init; } = [];
    public RcpPayTick? Selected { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class RcpMnrbSaveRequestValidator : Common.RequestValidator<RcpMnrbSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public RcpMnrbSaveRequestValidator()
    {
        RuleFor(r => r.ReceiptNo).MaximumLength(30);
        RuleFor(r => r.ChequeNo).MaximumLength(30);
        RuleFor(r => r.Remark).MaximumLength(500);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 5000 && t.All(x => x.BordxNo.Length <= 20 && (x.CasNumber ?? "").Length <= 30));
    }
}

/// <summary>Data access.</summary>
public interface IRcpMnrbRepository
{
#pragma warning disable CS1591
    Task<List<(string Code, string Name)>> PayorsAsync(IDbSession session);
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, RcpMnrbSearchRequest r, string? payor, DateTime? from, DateTime? to);
    Task<(string Status, List<string> Cases)> BordxCasesAsync(IDbSession session, string bordxNo);
    Task<string> CaseStatusAsync(IDbSession session, string casNumber, string bordxNo);
    Task SaveBordxCaseAsync(IDbSession session, string casNumber, string bordxNo, RcpPayReceipt r);
    Task SaveWksCaseAsync(IDbSession session, string casNumber, string bordxNo, string status, decimal selectedAmt, RcpPayReceipt r);
    Task<Dictionary<string, object?>?> ReceiptAsync(IDbSession session, string bordxNo, string? casNumber);
    Task EditAsync(IDbSession session, string bordxNo, string? casNumber, RcpPayReceipt r);
#pragma warning restore CS1591
}

/// <summary>Receipt > Fr MNRB (FrmReceiptMNRB).</summary>
public interface IRcpMnrbService
{
    /// <summary>The MNRB payors.</summary>
    Task<List<(string Code, string Name)>> PayorsAsync();

    /// <summary>cmdSearch_Click.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(RcpMnrbSearchRequest request);

    /// <summary>lstBordx_Click / lstWorksheet_Click in edit mode.</summary>
    Task<Dictionary<string, object?>> ReceiptAsync(string bordxNo, string? casNumber);

    /// <summary>cmdSave_Click (after the confirmation).</summary>
    Task<string> SaveAsync(RcpMnrbSaveRequest request, string user);

    /// <summary>cmdUpdate_Click (after the confirmation).</summary>
    Task<string> UpdateAsync(RcpMnrbSaveRequest request, string user);
}

/// <inheritdoc />
public sealed class RcpMnrbService(IDbSessions sessions, IRcpMnrbRepository repository) : IRcpMnrbService
{
    /// <summary>Desktop texts.</summary>
    public const string NoWorksheet = "Please Enter Worksheet No", NoBordx = "Please Enter Bordx No", NoRecord = "No record found!", NoRecordWks = "No record found !",
        NoneSelected = "Please select at least one record to be updated! ", NoReceiptNo = "Please Enter Receipt No!", NoReceiptNoWks = "Please Enter Receipt No !",
        NoReceiptDate = "Please Enter Receipt Date !", NoReceiptDateWks = "Please Enter Receipt Date", NoCheque = "Please Enter Cheque No!",
        NoChequeWks = "Please Enter Cheque No. !", NoAmount = "Please Enter Receipt Amount!", NoRemark = "Please Enter Remark!", Updated = "Record Updated...",
        UpdatedWks = "Record updated...", NoSelection = "Please select the record", BadDate = "Invalid Receipt Date!", NoAmountUpd = "Please Enter Receipt Amt!",
        Paid = "Record has been paid !", PartialPaid = "Record has been Partial Paid !", NotPaid = "Payment hasn't been made!";

    /// <inheritdoc />
    public async Task<List<(string Code, string Name)>> PayorsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.PayorsAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(RcpMnrbSearchRequest r)
    {
        if (r.Scan && r.Worksheet && string.IsNullOrWhiteSpace(r.CaseFrom)) throw new BusinessRuleException(NoWorksheet);
        if (r.Scan && !r.Worksheet && string.IsNullOrWhiteSpace(r.BordxFrom)) throw new BusinessRuleException(NoBordx);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, r, RcpPayService.PayorCode(r.Payor), ClinicInvoiceRules.Date(r.ReceiptDateFrom), ClinicInvoiceRules.Date(r.ReceiptDateTo));
        if (rows.Count == 0) throw new NotFoundException(r.Worksheet ? NoRecordWks : NoRecord);
        foreach (var x in rows) Shape(x, r.Worksheet);
        return rows;
    }

    /// <summary>ReceiptBordxListing / ReceiptWorksheetListing: status, balance and the amount a tick adds.</summary>
    public static void Shape(Dictionary<string, object?> x, bool worksheet)
    {
        var status = T(x, worksheet ? "PVRecMNRBStatus" : "BordxMNRBFRStatus");
        x["Status"] = status switch { "F" => "Full", "P" => "Partial", _ => "" };
        if (worksheet) x["TickAmount"] = N(x, "CLMTotalApproved");
        else
        {
            x["Balance"] = x.GetValueOrDefault("CheckAmt") is null or DBNull ? N(x, "BordxClaimAmt") : N(x, "BordxClaimAmt") - N(x, "CheckAmt");
            x["TickAmount"] = status == "P" ? x["Balance"] : N(x, "BordxClaimAmt");
        }
    }

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>> ReceiptAsync(string bordxNo, string? casNumber)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.ReceiptAsync(s, bordxNo.Trim(), Blank(casNumber)) ?? throw new BusinessRuleException(NotPaid);
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(RcpMnrbSaveRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(NoneSelected);
        if (string.IsNullOrWhiteSpace(r.ReceiptNo)) throw new BusinessRuleException(r.Worksheet ? NoReceiptNoWks : NoReceiptNo);
        var date = ClinicInvoiceRules.Date(r.ReceiptDate) ?? throw new BusinessRuleException(r.Worksheet ? NoReceiptDateWks : NoReceiptDate);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(r.Worksheet ? NoChequeWks : NoCheque);
        if (!r.Worksheet && string.IsNullOrWhiteSpace(r.ReceiptAmt)) throw new BusinessRuleException(NoAmount);
        var amount = Amount(r.ReceiptAmt);
        var receipt = new RcpPayReceipt(r.ReceiptNo.Trim(), date, amount, r.ChequeNo.Trim(), Blank(r.Remark), Blank(r.Month), Blank(r.Year), user);

        await using var tx = await sessions.BeginAsync();
        decimal selected = 0;
        var work = new List<(string Case, string Bordx)>();
        foreach (var t in r.Ticked.DistinctBy(t => (t.BordxNo.Trim(), (t.CasNumber ?? "").Trim())))
        {
            var bordx = t.BordxNo.Trim();
            if (r.Worksheet)
            {
                var cas = (t.CasNumber ?? "").Trim();
                if (await repository.CaseStatusAsync(tx, cas, bordx) == "F") throw new BusinessRuleException(Paid);
                var row = (await repository.SearchAsync(tx, new RcpMnrbSearchRequest { Worksheet = true, Scan = true, CaseFrom = cas, BordxFrom = bordx, BordxTo = bordx },
                    null, null, null)).FirstOrDefault();
                if (row is not null) { Shape(row, true); selected += (decimal)row["TickAmount"]!; }
                work.Add((cas, bordx));
            }
            else
            {
                var (status, cases) = await repository.BordxCasesAsync(tx, bordx);
                if (status == "F") throw new BusinessRuleException(Paid);
                var row = (await repository.SearchAsync(tx, new RcpMnrbSearchRequest { Scan = true, BordxFrom = bordx }, null, null, null)).FirstOrDefault();
                if (row is not null) { Shape(row, false); selected += (decimal)row["TickAmount"]!; }
                work.AddRange(cases.Select(c => (c, bordx)));
            }
        }
        if (amount != selected && string.IsNullOrWhiteSpace(r.Remark)) throw new BusinessRuleException(NoRemark);
        foreach (var (cas, bordx) in work)
            if (r.Worksheet) await repository.SaveWksCaseAsync(tx, cas, bordx, "F", selected, receipt);
            else await repository.SaveBordxCaseAsync(tx, cas, bordx, receipt);
        await tx.CommitAsync();
        return r.Worksheet ? UpdatedWks : Updated;
    }

    /// <inheritdoc />
    public async Task<string> UpdateAsync(RcpMnrbSaveRequest r, string user)
    {
        if (r.Selected is null) throw new BusinessRuleException(NoSelection);
        if (string.IsNullOrWhiteSpace(r.ReceiptNo)) throw new BusinessRuleException(NoReceiptNo);
        var date = ClinicInvoiceRules.Date(r.ReceiptDate)
            ?? throw new BusinessRuleException(string.IsNullOrWhiteSpace(r.ReceiptDate) ? NoReceiptDateWks : BadDate);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(NoCheque);
        if (string.IsNullOrWhiteSpace(r.ReceiptAmt)) throw new BusinessRuleException(NoAmountUpd);
        await using var tx = await sessions.BeginAsync();
        await repository.EditAsync(tx, r.Selected.BordxNo.Trim(), r.Worksheet ? Blank(r.Selected.CasNumber) : null,
            new RcpPayReceipt(r.ReceiptNo.Trim(), date, Amount(r.ReceiptAmt), r.ChequeNo.Trim(), Blank(r.Remark), Blank(r.Month), Blank(r.Year), user));
        await tx.CommitAsync();
        return Updated;
    }

    private static decimal Amount(string? v) => decimal.TryParse((v ?? "").Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : 0;

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
