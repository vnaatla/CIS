using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.AcExport;

/// <summary>Invoice Posting > Clinic search (frmPostingInvoiceCln SearchRecord).</summary>
public sealed record PostInvClnRequest
{
#pragma warning disable CS1591
    /// <summary>G: the debit and credit lines, D: debit only, C: credit only.</summary>
    public string? Kind { get; init; } = "G";
    /// <summary>"code - name".</summary>
    public string? Clinic { get; init; }
    public string? GroupCompany { get; init; }
    public string? InvDateFrom { get; init; }
    public string? InvDateTo { get; init; }
    public string? InvoiceFrom { get; init; }
    public string? InvoiceTo { get; init; }
    /// <summary>"0 - No" / "1 - Yes".</summary>
    public string? Posted { get; init; } = "0 - No";
    public string? PostDateFrom { get; init; }
    public string? PostDateTo { get; init; }
    public string? BatchFrom { get; init; }
    public string? BatchTo { get; init; }
    public string? BordxDateFrom { get; init; }
    public string? BordxDateTo { get; init; }
    /// <summary>The A/C period searched ("MM/YYYY").</summary>
    public string? InvoicePeriod { get; init; }
    public string? PvDateFrom { get; init; }
    public string? PvDateTo { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PostInvClnRequestValidator : Common.RequestValidator<PostInvClnRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostInvClnRequestValidator()
    {
        RuleFor(r => r.Clinic).MaximumLength(300);
        RuleFor(r => r.GroupCompany).MaximumLength(200);
        RuleFor(r => r.InvoiceFrom).MaximumLength(100);
        RuleFor(r => r.InvoiceTo).MaximumLength(100);
        RuleFor(r => r.BatchFrom).MaximumLength(20);
        RuleFor(r => r.BatchTo).MaximumLength(20);
        RuleFor(r => r.InvoicePeriod).MaximumLength(7);
    }
}

/// <summary>A ticked invoice.</summary>
public sealed record PostInvClnTick(string InvoiceNo, string ClinicCode, DateTime? InvoiceDate);

/// <summary>Post: the search (for the bordx dates of the file name), the period and the ticked invoices.</summary>
public sealed record PostInvClnPostRequest(PostInvClnRequest Search, string? Period, List<PostInvClnTick> Ticked);

/// <summary>Lengths.</summary>
public sealed class PostInvClnPostRequestValidator : Common.RequestValidator<PostInvClnPostRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostInvClnPostRequestValidator()
    {
        RuleFor(r => r.Period).MaximumLength(7);
        RuleFor(r => r.Search).SetValidator(new PostInvClnRequestValidator());
        RuleFor(r => r.Ticked).Must(t => t.Count <= 20000 && t.All(x => x.InvoiceNo.Length <= 100 && x.ClinicCode.Length <= 10));
    }
}

/// <summary>A posting line (lstClnInvList / lstCDList).</summary>
public sealed record PostingLine(bool Posted, string Journal, string Ledger, string BatchNo, string Period, string Date, string Reference, string Account, decimal Amount,
    string DebitCredit, string Analysis, string PayCode, string ClinicCode, bool PostLine, DateTime? InvoiceDate, string Description);

/// <summary>The lines, the debit / credit totals, the record count and whether the list is posted.</summary>
public sealed record PostingResult(List<PostingLine> Lines, decimal Debit, decimal Credit, int Records, bool Posted, string? BatchNo = null);

/// <summary>Data access.</summary>
public interface IPostInvClnRepository
{
#pragma warning disable CS1591
    Task<List<string>> PeriodsAsync(IDbSession session);
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostInvClnFilter filter);
    Task<string> NextInvNoAsync(IDbSession session);
    Task InsertAsync(IDbSession session, PostInvClnTick tick, string accMonth, string accYear, string batchNo, string user);
#pragma warning restore CS1591
}

/// <summary>The parsed criteria.</summary>
public sealed record PostInvClnFilter(string? BatchNo, string? Clinic, DateTime? InvDateFrom, DateTime? InvDateTo, string? InvoiceFrom, string? InvoiceTo, string Posted,
    DateTime? PostDateFrom, DateTime? PostDateTo, string? BatchFrom, string? BatchTo, string? BordxMonthFrom, string? BordxMonthTo, string? BordxYearFrom,
    string? BordxYearTo, string? AccMonth, string? AccYear, DateTime? PvDateFrom, DateTime? PvDateTo);

/// <summary>A/C Export > Invoice Posting > Clinic (frmPostingInvoiceCln).</summary>
public interface IPostInvClnService
{
    /// <summary>The accounting periods.</summary>
    Task<List<string>> PeriodsAsync();

    /// <summary>Search / Search Debit / Search Credit.</summary>
    Task<PostingResult> SearchAsync(PostInvClnRequest request);

    /// <summary>CmdPost_Click (after the confirmation): the batch no., its lines and the posting workbook.</summary>
    Task<(PostingResult Result, FileOutput File)> PostAsync(PostInvClnPostRequest request, string user);

    /// <summary>CmdPrintFile_Click: the list as a workbook.</summary>
    FileOutput Export(PostingResult result, string fileName);
}

/// <inheritdoc />
public sealed class PostInvClnService(IDbSessions sessions, IPostInvClnRepository repository, IReportExporter exporter) : IPostInvClnService
{
    /// <summary>Desktop texts.</summary>
    public const string NoSelection = "Please select the record", Posted = "Record(s) has been posted ! ", NoBordxDate = "Please Enter Bordx Date! ",
        NoPeriod = "Please Select Accounting Period ! ", Done = "Record(s) Posted Sucessfully... ! ", NoRecord = "No record selected ! ", BadDate = "Invalid Date Format! ",
        NoPrint = "Report cannot be printed without any record.";

    /// <inheritdoc />
    public async Task<List<string>> PeriodsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.PeriodsAsync(s);
    }

    /// <inheritdoc />
    public async Task<PostingResult> SearchAsync(PostInvClnRequest r)
    {
        var filter = Filter(r, null);
        await using var s = await sessions.OpenAsync();
        var result = Lines(await repository.SearchAsync(s, filter), Kind(r.Kind));
        if (result.Records == 0) throw new NotFoundException(NoRecord);
        return result;
    }

    /// <inheritdoc />
    public async Task<(PostingResult Result, FileOutput File)> PostAsync(PostInvClnPostRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(NoSelection);
        var bordxFrom = Date(r.Search.BordxDateFrom);
        var bordxTo = Date(r.Search.BordxDateTo);
        if (bordxFrom is null || bordxTo is null) throw new BusinessRuleException(NoBordxDate);
        var period = (r.Period ?? "").Trim();
        if (period.Length < 7) throw new BusinessRuleException(NoPeriod);
        await using var tx = await sessions.BeginAsync();
        // The ticked invoices must still be unposted (re-read with the same criteria).
        var open = (await repository.SearchAsync(tx, Filter(r.Search with { Posted = "0 - No" }, null)))
            .Select(x => (T(x, "InvoiceNo"), T(x, "ClinicCode"))).ToHashSet();
        if (r.Ticked.Any(t => !open.Contains((t.InvoiceNo.Trim(), t.ClinicCode.Trim())))) throw new BusinessRuleException(Posted);
        var batch = await repository.NextInvNoAsync(tx);
        foreach (var t in r.Ticked.DistinctBy(t => (t.InvoiceNo.Trim(), t.ClinicCode.Trim())))
            await repository.InsertAsync(tx, t with { InvoiceNo = t.InvoiceNo.Trim(), ClinicCode = t.ClinicCode.Trim() }, period[..2], period.Substring(3, 4), batch, user);
        var result = Lines(await repository.SearchAsync(tx, new PostInvClnFilter(batch, null, null, null, null, null, "", null, null, null, null, null, null, null, null, null,
            null, null, null)), "G") with { BatchNo = batch };
        await tx.CommitAsync();
        var name = $"{batch} RC-CLNInv {bordxFrom:yyyyMMdd}-{bordxTo:dd}";
        return (result, Export(result, name));
    }

    /// <inheritdoc />
    public FileOutput Export(PostingResult result, string fileName)
    {
        if (result.Lines.Count == 0) throw new BusinessRuleException(NoPrint);
        var rows = result.Lines.Select(l => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
        {
            ["Post"] = l.Posted ? "Post" : "", ["Journal"] = l.Journal, ["Ledger"] = l.Ledger, ["Batch"] = l.BatchNo, ["Period"] = l.Period, ["Date"] = l.Date,
            ["Reference"] = l.Reference, ["Account"] = l.Account, ["Amount"] = l.Amount, ["DC"] = l.DebitCredit, ["Analysis"] = l.Analysis, ["PayCode"] = l.PayCode,
            ["Clinic"] = l.ClinicCode, ["Description"] = l.Description,
        }).ToList();
        return exporter.Export(new ReportDocument("POSTING",
        [
            new("Post", "Post"), new("Journal", "Journal"), new("Ledger", "Ledger"), new("Batch", "Batch No"), new("Period", "Period"), new("Date", "Date"),
            new("Reference", "Reference"), new("Account", "Account"), new("Amount", "Amount", ReportFormat.Money), new("DC", "D/C"), new("Analysis", "Analysis"),
            new("PayCode", "Payor"), new("Clinic", "Clinic"), new("Description", "Description"),
        ], rows) { FileName = fileName }, ReportFileType.Excel);
    }

    /// <summary>ClnInvListing: G gives a credit line (84 + clinic, the clinic name) and a debit line (6000, YYMM-clinic name, the line Post writes) per invoice.</summary>
    public static PostingResult Lines(List<Dictionary<string, object?>> rows, string kind)
    {
        var lines = new List<PostingLine>();
        decimal debit = 0, credit = 0;
        var posted = false;
        foreach (var x in rows)
        {
            var isPosted = T(x, "InvCLNStatus") == "1";
            posted |= isPosted;
            var amount = Math.Round(N(x, "CLMPayableByMedix2C"), 2);
            var inv = x.GetValueOrDefault("InvoiceDate") as DateTime?;
            var period = T(x, "AccMonth") + T(x, "AccYear") == "" ? "" : $"{T(x, "AccMonth")}/{T(x, "AccYear")}";
            var date = inv?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) ?? "";
            var reference = T(x, "InvoiceNo");
            if (reference.Length > 100) reference = reference[..100];
            PostingLine Line(string account, string dc, string desc, bool postLine) => new(isPosted, "RC", "UHC", T(x, "BatchNo"), period, date, reference, account, amount, dc,
                "4" + T(x, "ClinicCode"), T(x, "PAYCode"), T(x, "ClinicCode"), postLine, inv, desc);
            var debitDesc = inv is null ? T(x, "ClnName") : $"{inv:yyMM}-{T(x, "ClnName")}";
            if (kind is "G" or "C") { lines.Add(Line("84" + T(x, "ClinicCode"), "C", T(x, "ClnName"), false)); credit += amount; }
            if (kind is "G" or "D") { lines.Add(Line("6000", "D", debitDesc, kind == "G")); debit += amount; }
        }
        return new PostingResult(lines, debit, credit, rows.Count, posted);
    }

    private PostInvClnFilter Filter(PostInvClnRequest r, string? batch)
    {
        var clinic = (r.Clinic ?? "").Trim();
        var i = clinic.IndexOf('-');
        var bordxFrom = Date(r.BordxDateFrom);
        var bordxTo = Date(r.BordxDateTo);
        var period = (r.InvoicePeriod ?? "").Trim();
        return new PostInvClnFilter(batch, clinic.Length == 0 ? null : (i > 0 ? clinic[..i] : clinic).Trim(), Date(r.InvDateFrom), Date(r.InvDateTo), Blank(r.InvoiceFrom),
            Blank(r.InvoiceTo), ((r.Posted ?? "0").Trim() + "0")[..1], Date(r.PostDateFrom), Date(r.PostDateTo), Blank(r.BatchFrom), Blank(r.BatchTo),
            bordxFrom?.ToString("MM", CultureInfo.InvariantCulture), bordxTo?.ToString("MM", CultureInfo.InvariantCulture),
            bordxFrom?.ToString("yyyy", CultureInfo.InvariantCulture), bordxTo?.ToString("yyyy", CultureInfo.InvariantCulture),
            period.Length >= 7 ? period[..2] : null, period.Length >= 7 ? period.Substring(3, 4) : null, Date(r.PvDateFrom), Date(r.PvDateTo));
    }

    private static string Kind(string? k) => (k ?? "G").Trim().ToUpperInvariant() switch { "D" => "D", "C" => "C", _ => "G" };

    private static DateTime? Date(string? v) =>
        string.IsNullOrWhiteSpace(v) ? null : ClinicInvoiceRules.Date(v) ?? throw new BusinessRuleException(BadDate);

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
