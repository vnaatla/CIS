using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Receipt;

/// <summary>Receipt from Group Company search (FrmReceiptPayor SearchRecord).</summary>
public sealed record RcpPaySearchRequest
{
#pragma warning disable CS1591
    /// <summary>false: by bordereaux (OptBordx), true: by worksheet (OptWks).</summary>
    public bool Worksheet { get; init; }
    public bool Scan { get; init; } = true;
    public string? Company { get; init; }
    public string? Subsidiary { get; init; }
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
    /// <summary>"code - name" (CISDB..PAY).</summary>
    public string? Payor { get; init; }
    public string? InsuredClient { get; init; }
    public string? ClientPvFrom { get; init; }
    public string? ClientPvTo { get; init; }
    public string? MedixReceiptFrom { get; init; }
    public string? MedixReceiptTo { get; init; }
    /// <summary>"", YES or NO.</summary>
    public string? Insured { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class RcpPaySearchRequestValidator : Common.RequestValidator<RcpPaySearchRequest>
{
    /// <summary>Defines the rules.</summary>
    public RcpPaySearchRequestValidator()
    {
        RuleFor(r => r.Company).MaximumLength(200);
        RuleFor(r => r.Subsidiary).MaximumLength(200);
        RuleFor(r => r.Payor).MaximumLength(300);
        RuleFor(r => r.InsuredClient).MaximumLength(200);
        RuleFor(r => r.BordxFrom).MaximumLength(20);
        RuleFor(r => r.BordxTo).MaximumLength(20);
        RuleFor(r => r.CaseFrom).MaximumLength(30);
        RuleFor(r => r.CaseTo).MaximumLength(30);
        RuleFor(r => r.Insured).MaximumLength(3);
    }
}

/// <summary>A ticked row: the bordereaux (bordereaux mode) or the case and its bordereaux (worksheet mode).</summary>
public sealed record RcpPayTick(string BordxNo, string? CasNumber);

/// <summary>Save (the ticked rows) / Update (the selected row).</summary>
public sealed record RcpPaySaveRequest
{
#pragma warning disable CS1591
    public bool Worksheet { get; init; }
    public string? Company { get; init; }
    public string? ReceiptNo { get; init; }
    public string? ReceiptDate { get; init; }
    public string? ChequeNo { get; init; }
    public string? ReceiptAmt { get; init; }
    public string? Remark { get; init; }
    public string? Month { get; init; }
    public string? Year { get; init; }
    /// <summary>The OP claim / OP fee lines (bordereaux mode).</summary>
    public string? OpClaimAmt { get; init; }
    public string? OpClaimDesc { get; init; }
    public string? OpFeeAmt { get; init; }
    public string? OpFeeDesc { get; init; }
    public List<RcpPayTick> Ticked { get; init; } = [];
    /// <summary>Update: the selected row.</summary>
    public RcpPayTick? Selected { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class RcpPaySaveRequestValidator : Common.RequestValidator<RcpPaySaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public RcpPaySaveRequestValidator()
    {
        RuleFor(r => r.ReceiptNo).MaximumLength(30);
        RuleFor(r => r.ChequeNo).MaximumLength(30);
        RuleFor(r => r.Remark).MaximumLength(500);
        RuleFor(r => r.OpClaimDesc).MaximumLength(500);
        RuleFor(r => r.OpFeeDesc).MaximumLength(500);
        RuleFor(r => r.Company).MaximumLength(200);
        RuleFor(r => r.Ticked).Must(t => t.Count <= 5000 && t.All(x => x.BordxNo.Length <= 20 && (x.CasNumber ?? "").Length <= 30));
    }
}

/// <summary>A receipt line written by Save / Update.</summary>
public record RcpPayReceipt(string ReceiptNo, DateTime ReceiptDate, decimal ReceiptAmt, string ChequeNo, string? Remark, string? Month, string? Year, string User);

/// <summary>Data access.</summary>
public interface IRcpPayRepository
{
#pragma warning disable CS1591
    Task<(List<string> Companies, List<(string Code, string Name)> Payors)> LookupsAsync(IDbSession session);
    Task<List<string>> ClientsAsync(IDbSession session, string payor);
    Task<List<string>> SubsidiariesAsync(IDbSession session, string company);
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, RcpPaySearchRequest r, string? payor, DateTime? from, DateTime? to, string? receiptNo);
    Task<(string Status, List<string> Cases)> BordxCasesAsync(IDbSession session, string bordxNo);
    Task<string> CaseStatusAsync(IDbSession session, string casNumber, string bordxNo);
    Task SaveBordxCaseAsync(IDbSession session, string casNumber, string bordxNo, RcpPayReceipt r);
    Task SaveWksCaseAsync(IDbSession session, string casNumber, string bordxNo, string status, RcpPayReceipt r);
    Task SummaryAsync(IDbSession session, string receiptNo, DateTime date, decimal amount, string type, string? desc, string chequeNo, string payCode);
    Task<Dictionary<string, object?>?> ReceiptAsync(IDbSession session, string bordxNo, string? casNumber);
    Task EditAsync(IDbSession session, string bordxNo, string? casNumber, RcpPayReceipt r);
#pragma warning restore CS1591
}

/// <summary>Receipt > Fr Grp Company (FrmReceiptPayor).</summary>
public interface IRcpPayService
{
    /// <summary>The combos.</summary>
    Task<object> LookupsAsync();

    /// <summary>cisdbcbopayor_LostFocus: the payor's insured clients.</summary>
    Task<List<string>> ClientsAsync(string payor);

    /// <summary>CboPayor_LostFocus: the group company's subsidiaries.</summary>
    Task<List<string>> SubsidiariesAsync(string company);

    /// <summary>cmdSearch_Click (or the receipt just saved).</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(RcpPaySearchRequest request, string? receiptNo = null);

    /// <summary>lstBordx_Click / lstWorksheet_Click in edit mode.</summary>
    Task<Dictionary<string, object?>> ReceiptAsync(string bordxNo, string? casNumber);

    /// <summary>CmdSave_Click (after the confirmation).</summary>
    Task<string> SaveAsync(RcpPaySaveRequest request, string user);

    /// <summary>CmdUpdate_Click (after the confirmation).</summary>
    Task<string> UpdateAsync(RcpPaySaveRequest request, string user);
}

/// <inheritdoc />
public sealed class RcpPayService(IDbSessions sessions, IRcpPayRepository repository) : IRcpPayService
{
    /// <summary>Desktop texts.</summary>
    public const string NoWorksheet = "Please Enter Worksheet No", NoBordx = "Please Enter Bordx No", NoRecord = "No record found!", NoRecordWks = "No record found !",
        NoneSelected = "Please select at least one record to be updated! ", NoCompany = "Please select group company! ", NoCompanyFee = "Please select group company!",
        NoReceiptNo = "Please Enter Receipt No!", NoReceiptNoWks = "Please Enter Receipt No !", NoReceiptDate = "Please Enter Receipt Date!",
        NoReceiptDateWks = "Please Enter Receipt Date", NoCheque = "Please Enter Cheque No!", NoChequeWks = "Please Enter Cheque No. !",
        NoAmount = "Please Enter Receipt Amount!", NoRemark = "Please Enter Remark!", Updated = "Record Updated...", UpdatedWks = "Record updated...",
        NoSelection = "Please select the record", BadDate = "Invalid Receipt Date!", NoAmountUpd = "Please Enter Receipt Amt!", Paid = "Record has been paid !",
        PartialPaid = "Record has been Partial Paid !", NotPaid = "Payment hasn't been made!";

    /// <inheritdoc />
    public async Task<object> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        var (companies, payors) = await repository.LookupsAsync(s);
        return new { companies, payors = payors.Select(p => new { code = p.Code, name = p.Name }) };
    }

    /// <inheritdoc />
    public async Task<List<string>> ClientsAsync(string payor)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.ClientsAsync(s, PayorCode(payor) ?? "");
    }

    /// <inheritdoc />
    public async Task<List<string>> SubsidiariesAsync(string company)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.SubsidiariesAsync(s, company.Trim());
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(RcpPaySearchRequest r, string? receiptNo = null)
    {
        if (receiptNo is null && r.Scan && r.Worksheet && string.IsNullOrWhiteSpace(r.CaseFrom)) throw new BusinessRuleException(NoWorksheet);
        if (receiptNo is null && r.Scan && !r.Worksheet && string.IsNullOrWhiteSpace(r.BordxFrom)) throw new BusinessRuleException(NoBordx);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, r, PayorCode(r.Payor), ClinicInvoiceRules.Date(r.ReceiptDateFrom), ClinicInvoiceRules.Date(r.ReceiptDateTo), receiptNo);
        if (rows.Count == 0) throw new NotFoundException(r.Worksheet ? NoRecordWks : NoRecord);
        foreach (var x in rows) Shape(x, r.Worksheet);
        return rows;
    }

    /// <summary>ReceiptBordxListing / ReceiptWorksheetListing: status, balance and the amount a tick adds.</summary>
    public static void Shape(Dictionary<string, object?> x, bool worksheet)
    {
        if (worksheet)
        {
            var status = T(x, "PVRecPayorStatus");
            x["Status"] = status switch { "F" => "Full", "P" => "Partial", _ => "" };
            x["Balance"] = N(x, "CLMBordxAmt") - N(x, "CheckAmt");
            x["TickError"] = status == "F" ? Paid : null;
            x["TickAmount"] = N(x, "CLMBordxAmt");
        }
        else
        {
            var status = T(x, "BordxPAYFRStatus").Length > 0 && T(x, "ReceiptNo").Length > 0 ? T(x, "BordxPAYFRStatus") : "";
            x["Status"] = status switch { "F" => "Full", "P" => "Partial", _ => "" };
            x["Balance"] = x.GetValueOrDefault("CheckAmt") is null or DBNull ? N(x, "BordxClaimAmt") : N(x, "BordxClaimAmt") - N(x, "CheckAmt");
            x["TickAmount"] = N(x, "BordxClaimAmt");
        }
    }

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>> ReceiptAsync(string bordxNo, string? casNumber)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.ReceiptAsync(s, bordxNo.Trim(), Blank(casNumber)) ?? throw new BusinessRuleException(NotPaid);
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(RcpPaySaveRequest r, string user)
    {
        var opClaim = Amount(r.OpClaimAmt) ?? 0;
        var opFee = Amount(r.OpFeeAmt) ?? 0;
        var company = (r.Company ?? "").Trim();
        if (r.Worksheet)
        {
            if (r.Ticked.Count == 0) throw new BusinessRuleException(NoneSelected);
            if (string.IsNullOrWhiteSpace(r.ReceiptNo)) throw new BusinessRuleException(NoReceiptNoWks);
        }
        else
        {
            if (r.Ticked.Count == 0 && opClaim != 0) throw new BusinessRuleException(NoneSelected);
            if (company.Length == 0 && opClaim != 0) throw new BusinessRuleException(NoCompany);
            if (company.Length == 0 && opFee != 0) throw new BusinessRuleException(NoCompanyFee);
            if (string.IsNullOrWhiteSpace(r.ReceiptNo)) throw new BusinessRuleException(NoReceiptNo);
        }
        var date = ClinicInvoiceRules.Date(r.ReceiptDate) ?? throw new BusinessRuleException(r.Worksheet ? NoReceiptDateWks : NoReceiptDate);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(r.Worksheet ? NoChequeWks : NoCheque);
        if (!r.Worksheet && string.IsNullOrWhiteSpace(r.ReceiptAmt)) throw new BusinessRuleException(NoAmount);
        var amount = Amount(r.ReceiptAmt) ?? 0;
        var receipt = new RcpPayReceipt(r.ReceiptNo!.Trim(), date, amount, r.ChequeNo.Trim(), Blank(r.Remark), Blank(r.Month), Blank(r.Year), user);

        await using var tx = await sessions.BeginAsync();
        // The selected amount is recomputed from the list as it is now (the desktop added the ticked rows' list amounts).
        decimal selected = 0;
        var work = new List<(string Case, string Bordx)>();
        foreach (var t in r.Ticked.DistinctBy(t => (t.BordxNo.Trim(), (t.CasNumber ?? "").Trim())))
        {
            var bordx = t.BordxNo.Trim();
            if (r.Worksheet)
            {
                var cas = (t.CasNumber ?? "").Trim();
                if (await repository.CaseStatusAsync(tx, cas, bordx) == "F") throw new BusinessRuleException(Paid);
                var row = (await repository.SearchAsync(tx, new RcpPaySearchRequest { Worksheet = true, Scan = true, CaseFrom = cas, BordxFrom = bordx, BordxTo = bordx },
                    null, null, null, null)).FirstOrDefault();
                if (row is not null) { Shape(row, true); selected += (decimal)row["TickAmount"]!; }
                work.Add((cas, bordx));
            }
            else
            {
                var (status, cases) = await repository.BordxCasesAsync(tx, bordx);
                if (status == "F") throw new BusinessRuleException(Paid);
                var row = (await repository.SearchAsync(tx, new RcpPaySearchRequest { Scan = true, BordxFrom = bordx }, null, null, null, null)).FirstOrDefault();
                if (row is not null) { Shape(row, false); selected += status == "P" ? (decimal)row["Balance"]! : (decimal)row["TickAmount"]!; }
                work.AddRange(cases.Select(c => (c, bordx)));
            }
        }
        if (amount != selected && string.IsNullOrWhiteSpace(r.Remark)) throw new BusinessRuleException(NoRemark);
        foreach (var (cas, bordx) in work)
            if (r.Worksheet) await repository.SaveWksCaseAsync(tx, cas, bordx, "F", receipt);
            else await repository.SaveBordxCaseAsync(tx, cas, bordx, receipt);
        if (!r.Worksheet)
        {
            if (opClaim != 0) await repository.SummaryAsync(tx, receipt.ReceiptNo, date, opClaim, "OPCLM", Blank(r.OpClaimDesc), receipt.ChequeNo, company);
            if (opFee != 0) await repository.SummaryAsync(tx, receipt.ReceiptNo, date, opFee, "OPFEE", Blank(r.OpFeeDesc), receipt.ChequeNo, company);
        }
        await tx.CommitAsync();
        return r.Worksheet ? UpdatedWks : Updated;
    }

    /// <inheritdoc />
    public async Task<string> UpdateAsync(RcpPaySaveRequest r, string user)
    {
        if (r.Selected is null) throw new BusinessRuleException(NoSelection);
        if (string.IsNullOrWhiteSpace(r.ReceiptNo)) throw new BusinessRuleException(NoReceiptNo);
        var date = ClinicInvoiceRules.Date(r.ReceiptDate)
            ?? throw new BusinessRuleException(string.IsNullOrWhiteSpace(r.ReceiptDate) ? NoReceiptDateWks : BadDate);
        if (string.IsNullOrWhiteSpace(r.ChequeNo)) throw new BusinessRuleException(NoCheque);
        if (string.IsNullOrWhiteSpace(r.ReceiptAmt)) throw new BusinessRuleException(NoAmountUpd);
        await using var tx = await sessions.BeginAsync();
        await repository.EditAsync(tx, r.Selected.BordxNo.Trim(), r.Worksheet ? Blank(r.Selected.CasNumber) : null,
            new RcpPayReceipt(r.ReceiptNo.Trim(), date, Amount(r.ReceiptAmt) ?? 0, r.ChequeNo.Trim(), Blank(r.Remark), Blank(r.Month), Blank(r.Year), user));
        await tx.CommitAsync();
        return Updated;
    }

    /// <summary>The payor code of "code - name" (the first 4 characters without a dash, as the desktop).</summary>
    public static string? PayorCode(string? v)
    {
        var t = (v ?? "").Trim();
        if (t.Length == 0) return null;
        var i = t.IndexOf('-');
        return (i < 0 ? t[..Math.Min(4, t.Length)] : t[..i]).Trim();
    }

    private static decimal? Amount(string? v) =>
        decimal.TryParse((v ?? "").Replace(",", ""), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : null;

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
