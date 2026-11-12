using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.AcExport;

/// <summary>Receipt Posting > Grp Company search (frmPostingRecGrpCompany SearchRecord).</summary>
public sealed record PostRecGrpRequest
{
#pragma warning disable CS1591
    /// <summary>G: all lines, D: debit only, C: credit only.</summary>
    public string? Kind { get; init; } = "G";
    public string? GroupCompany { get; init; }
    public string? ReceiptFrom { get; init; }
    public string? ReceiptTo { get; init; }
    public string? ReceiptDateFrom { get; init; }
    public string? ReceiptDateTo { get; init; }
    public string? PostDateFrom { get; init; }
    public string? PostDateTo { get; init; }
    /// <summary>"0 - No" / "1 - Yes".</summary>
    public string? Posted { get; init; } = "0 - No";
    public string? BatchFrom { get; init; }
    public string? BatchTo { get; init; }
    /// <summary>0 Health Connect (payors other than SJ), 1 Sihat Johor (SJ).</summary>
    public int Entity { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PostRecGrpRequestValidator : Common.RequestValidator<PostRecGrpRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostRecGrpRequestValidator()
    {
        RuleFor(r => r.GroupCompany).MaximumLength(200);
        RuleFor(r => r.ReceiptFrom).MaximumLength(30);
        RuleFor(r => r.ReceiptTo).MaximumLength(30);
        RuleFor(r => r.BatchFrom).MaximumLength(20);
        RuleFor(r => r.BatchTo).MaximumLength(20);
    }
}

/// <summary>Post: the search, the accounting period and the ticked receipt nos.</summary>
public sealed record PostRecGrpPostRequest(PostRecGrpRequest Search, string? Period, List<string> Ticked);

/// <summary>Lengths.</summary>
public sealed class PostRecGrpPostRequestValidator : Common.RequestValidator<PostRecGrpPostRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostRecGrpPostRequestValidator()
    {
        RuleFor(r => r.Period).MaximumLength(7);
        RuleFor(r => r.Search).SetValidator(new PostRecGrpRequestValidator());
        RuleFor(r => r.Ticked).Must(t => t.Count <= 20000 && t.All(x => x.Length <= 30));
    }
}

/// <summary>The parsed criteria.</summary>
public sealed record PostRecGrpFilter(string? BatchNo, string? Company, string? ReceiptFrom, string? ReceiptTo, DateTime? ReceiptDateFrom, DateTime? ReceiptDateTo,
    DateTime? PostDateFrom, DateTime? PostDateTo, string Posted, string? BatchFrom, string? BatchTo, bool Sj);

/// <summary>Data access.</summary>
public interface IPostRecGrpRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostRecGrpFilter filter);
    Task InsertAsync(IDbSession session, string receiptNo, string accMonth, string accYear, string batchNo, string user);
#pragma warning restore CS1591
}

/// <summary>A/C Export > Receipt Posting > Grp Company (frmPostingRecGrpCompany).</summary>
public interface IPostRecGrpService
{
    /// <summary>Search / Search Debit / Search Credit.</summary>
    Task<PostingResult> SearchAsync(PostRecGrpRequest request);

    /// <summary>CmdExp_Click (after the confirmation): the batch no., its lines and the posting workbook.</summary>
    Task<(PostingResult Result, FileOutput File)> PostAsync(PostRecGrpPostRequest request, string user);
}

/// <inheritdoc />
public sealed class PostRecGrpService(IDbSessions sessions, IPostRecGrpRepository repository, IPostInvClnGstRepository numbers, IPostInvClnService export)
    : IPostRecGrpService
{
    /// <summary>Desktop text (the others are those of <see cref="PostInvClnService"/>).</summary>
    public const string NoReceiptDate = "Please Enter Receipt Date! ";

    /// <inheritdoc />
    public async Task<PostingResult> SearchAsync(PostRecGrpRequest r)
    {
        await using var s = await sessions.OpenAsync();
        var result = Lines(await repository.SearchAsync(s, Filter(r, null)), Kind(r.Kind));
        if (result.Records == 0) throw new NotFoundException(PostBordxGrpService.NoRecord);
        return result;
    }

    /// <inheritdoc />
    public async Task<(PostingResult Result, FileOutput File)> PostAsync(PostRecGrpPostRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(PostInvClnService.NoSelection);
        var from = Date(r.Search.ReceiptDateFrom);
        var to = Date(r.Search.ReceiptDateTo);
        if (from is null || to is null) throw new BusinessRuleException(NoReceiptDate);
        var period = (r.Period ?? "").Trim();
        if (period.Length < 7) throw new BusinessRuleException(PostPayClnService.NoPeriod);
        await using var tx = await sessions.BeginAsync();
        var open = (await repository.SearchAsync(tx, Filter(r.Search with { Posted = "0 - No" }, null))).Select(x => T(x, "ReceiptNo"))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (r.Ticked.Any(t => !open.Contains(t.Trim()))) throw new BusinessRuleException(PostInvClnService.Posted);
        var batch = await numbers.NextBatchAsync(tx);
        foreach (var t in r.Ticked.Select(t => t.Trim()).Distinct())
            await repository.InsertAsync(tx, t, period[..2], period.Substring(3, 4), batch, user);
        var result = Lines(await repository.SearchAsync(tx, new PostRecGrpFilter(batch, null, null, null, null, null, null, null, "", null, null, r.Search.Entity == 1)),
            "G") with { BatchNo = batch };
        await tx.CommitAsync();
        return (result, export.Export(result, $"{batch} CR-GrpCompany {from:yyyyMMdd}-{to:dd}"));
    }

    /// <summary>GrpCompListing: per receipt a debit 7610 (as Search Debit) and a credit on the company account (65 + group company for KDC bordereaux, else 65 + the
    /// payor's account code characters 3–6, else 65 + group company); the credit line is the one Post writes. The desktop put both lines on the company account and
    /// listed (and totalled) a receipt once per bordereaux; here each receipt is listed once with its bordereaux. Search Credit lists 65 + group company, Search
    /// Debit 7610.</summary>
    public static PostingResult Lines(List<Dictionary<string, object?>> rows, string kind)
    {
        var lines = new List<PostingLine>();
        decimal total = 0;
        var posted = false;
        foreach (var g in rows.GroupBy(x => T(x, "ReceiptNo")))
        {
            var x = g.FirstOrDefault(r => T(r, "BordxRefNo").StartsWith("KDC", StringComparison.Ordinal)) ?? g.First();
            var isPosted = T(x, "RecGrpCompStatus") == "1";
            posted |= isPosted;
            var bordx = T(x, "BordxRefNo");
            var date = x.GetValueOrDefault("ReceiptDate") as DateTime?;
            var amount = N(x, "CheckAmt");
            var acc = T(x, "PAYAccCode");
            var account = bordx.StartsWith("KDC", StringComparison.Ordinal) || acc.Length < 3 ? "65" + T(x, "GRPCoCode")
                : "65" + acc.Substring(2, Math.Min(4, acc.Length - 2));
            var allBordx = string.Join(", ", g.Select(r => T(r, "BordxRefNo")).Where(b => b.Length > 0).Distinct());
            PostingLine Line(string a, string dc, bool postLine) => new(isPosted, "CR", "UHC", T(x, "BatchNo"),
                T(x, "AccMonth") + T(x, "AccYear") == "" ? "" : $"{T(x, "AccMonth")}/{T(x, "AccYear")}", date?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) ?? "",
                T(x, "ReceiptNo"), a, amount, dc, "", T(x, "PAYCode"), T(x, "GRPCoCode"), postLine, date, T(x, "MBMGroupCompany"))
            {
                Extra = new Dictionary<string, string> { ["Bordx"] = allBordx },
            };
            if (kind == "G")
            {
                lines.Add(Line("7610", "D", false));
                lines.Add(Line(account, "C", true));
            }
            else lines.Add(kind == "C" ? Line("65" + T(x, "GRPCoCode"), "C", false) : Line("7610", "D", false));
            total += amount;
        }
        return new PostingResult(lines, total, total, lines.Select(l => l.Reference).Distinct().Count(), posted);
    }

    private static PostRecGrpFilter Filter(PostRecGrpRequest r, string? batch)
    {
        var from = Date(r.ReceiptDateFrom);
        var to = Date(r.ReceiptDateTo);
        var postFrom = Date(r.PostDateFrom);
        var postTo = Date(r.PostDateTo);
        // Each date bound applies on its own (the desktop applied the receipt and post dates only when both were given).
        var posted = (r.Posted ?? "").Trim();
        return new PostRecGrpFilter(batch, Blank(r.GroupCompany), Blank(r.ReceiptFrom), Blank(r.ReceiptTo), from, to, postFrom, postTo,
            posted.Length > 0 ? posted[..1] : "", Blank(r.BatchFrom), Blank(r.BatchTo), r.Entity == 1);
    }

    private static string Kind(string? k) => (k ?? "G").Trim().ToUpperInvariant() switch { "D" => "D", "C" => "C", _ => "G" };

    private static DateTime? Date(string? v) =>
        string.IsNullOrWhiteSpace(v) ? null : ClinicInvoiceRules.Date(v) ?? throw new BusinessRuleException(PostInvClnService.BadDate);

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
