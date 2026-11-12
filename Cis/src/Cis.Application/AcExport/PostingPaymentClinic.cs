using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.AcExport;

/// <summary>Payment Posting > Clinic search (FrmPostingPay2Cln SearchRecord).</summary>
public sealed record PostPayClnRequest
{
#pragma warning disable CS1591
    /// <summary>G: all lines, D: debit only, C: credit only.</summary>
    public string? Kind { get; init; } = "G";
    /// <summary>"code - name".</summary>
    public string? Clinic { get; init; }
    public string? PvDateFrom { get; init; }
    public string? PvDateTo { get; init; }
    public string? PvFrom { get; init; }
    public string? PvTo { get; init; }
    /// <summary>"0 - No" / "1 - Yes".</summary>
    public string? Posted { get; init; } = "0 - No";
    public string? PostDateFrom { get; init; }
    public string? PostDateTo { get; init; }
    public string? BatchFrom { get; init; }
    public string? BatchTo { get; init; }
    /// <summary>0 Health Connect (payors other than SJ), 1 Sihat Johor (SJ).</summary>
    public int Entity { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PostPayClnRequestValidator : Common.RequestValidator<PostPayClnRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostPayClnRequestValidator()
    {
        RuleFor(r => r.Clinic).MaximumLength(300);
        RuleFor(r => r.PvFrom).MaximumLength(30);
        RuleFor(r => r.PvTo).MaximumLength(30);
        RuleFor(r => r.BatchFrom).MaximumLength(20);
        RuleFor(r => r.BatchTo).MaximumLength(20);
    }
}

/// <summary>A ticked PV (and its line's analysis code, which the desktop writes as InvoiceNo).</summary>
public sealed record PostPayClnTick(string PvNo, string Analysis);

/// <summary>Post: the search, the accounting period and the ticked PVs.</summary>
public sealed record PostPayClnPostRequest(PostPayClnRequest Search, string? Period, List<PostPayClnTick> Ticked);

/// <summary>Lengths.</summary>
public sealed class PostPayClnPostRequestValidator : Common.RequestValidator<PostPayClnPostRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostPayClnPostRequestValidator()
    {
        RuleFor(r => r.Period).MaximumLength(7);
        RuleFor(r => r.Search).SetValidator(new PostPayClnRequestValidator());
        RuleFor(r => r.Ticked).Must(t => t.Count <= 20000 && t.All(x => x.PvNo.Length <= 30 && x.Analysis.Length <= 30));
    }
}

/// <summary>The parsed criteria.</summary>
public sealed record PostPayClnFilter(string? BatchNo, string? Clinic, DateTime? PvDateFrom, DateTime? PvDateTo, string? PvFrom, string? PvTo, string Posted,
    DateTime? PostDateFrom, DateTime? PostDateTo, string? BatchFrom, string? BatchTo, bool Sj);

/// <summary>Data access.</summary>
public interface IPostPayClnRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostPayClnFilter filter);
    Task InsertAsync(IDbSession session, PostPayClnTick tick, string accMonth, string accYear, string batchNo, string user);
#pragma warning restore CS1591
}

/// <summary>A/C Export > Payment Posting > Clinic (FrmPostingPay2Cln).</summary>
public interface IPostPayClnService
{
    /// <summary>Search / Search Debit / Search Credit.</summary>
    Task<PostingResult> SearchAsync(PostPayClnRequest request);

    /// <summary>CmdPost_Click (after the confirmation): the batch no., its lines and the posting workbook.</summary>
    Task<(PostingResult Result, FileOutput File)> PostAsync(PostPayClnPostRequest request, string user);
}

/// <inheritdoc />
public sealed class PostPayClnService(IDbSessions sessions, IPostPayClnRepository repository, IPostInvClnGstRepository numbers, IPostInvClnService export)
    : IPostPayClnService
{
    /// <summary>Desktop texts (the others are those of <see cref="PostInvClnService"/>).</summary>
    public const string NoPvDate = "Please Enter PV Date! ", NoPeriod = "Please Keyin for Accounting Period ! ";

    /// <inheritdoc />
    public async Task<PostingResult> SearchAsync(PostPayClnRequest r)
    {
        await using var s = await sessions.OpenAsync();
        var result = Lines(await repository.SearchAsync(s, Filter(r, null)), Kind(r.Kind));
        if (result.Records == 0) throw new NotFoundException(PostInvClnService.NoRecord);
        return result;
    }

    /// <inheritdoc />
    public async Task<(PostingResult Result, FileOutput File)> PostAsync(PostPayClnPostRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(PostInvClnService.NoSelection);
        var pvFrom = Date(r.Search.PvDateFrom);
        var pvTo = Date(r.Search.PvDateTo);
        if (pvFrom is null || pvTo is null) throw new BusinessRuleException(NoPvDate);
        var period = (r.Period ?? "").Trim();
        if (period.Length < 7) throw new BusinessRuleException(NoPeriod);
        await using var tx = await sessions.BeginAsync();
        var open = (await repository.SearchAsync(tx, Filter(r.Search with { Posted = "0 - No" }, null))).Select(x => T(x, "PVCLNNo")).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (r.Ticked.Any(t => !open.Contains(t.PvNo.Trim()))) throw new BusinessRuleException(PostInvClnService.Posted);
        var batch = await numbers.NextBatchAsync(tx);
        foreach (var t in r.Ticked.DistinctBy(t => t.PvNo.Trim()))
            await repository.InsertAsync(tx, t with { PvNo = t.PvNo.Trim(), Analysis = t.Analysis.Trim() }, period[..2], period.Substring(3, 4), batch, user);
        var result = Lines(await repository.SearchAsync(tx, new PostPayClnFilter(batch, null, null, null, null, null, "", null, null, null, null, r.Search.Entity == 1)), "G")
            with { BatchNo = batch };
        await tx.CommitAsync();
        return (result, export.Export(result, $"{batch} CP-PayCLN {pvFrom:yyyyMMdd}-{pvTo:dd}"));
    }

    /// <summary>ClnPayListing: per PV a credit line (7520 for payor SJ, else 7610; the clinic name) and a debit line (84 + clinic, the line Post writes); journal CP,
    /// ledger MS for SJ, else UHC.</summary>
    public static PostingResult Lines(List<Dictionary<string, object?>> rows, string kind)
    {
        var lines = new List<PostingLine>();
        decimal debit = 0, credit = 0;
        var posted = false;
        foreach (var x in rows)
        {
            var isPosted = T(x, "Pay2ClnStatus") == "1";
            posted |= isPosted;
            var sj = T(x, "PayCode") == "SJ";
            var amount = Math.Round(N(x, "CLNPaidAmt"), 2, MidpointRounding.AwayFromZero);
            var pv = x.GetValueOrDefault("PVCLNDate") as DateTime?;
            var period = T(x, "AccMonth") + T(x, "AccYear") == "" ? "" : $"{T(x, "AccMonth")}/{T(x, "AccYear")}";
            var desc = T(x, "CLNName");
            if (desc.Length > 25) desc = desc[..25];
            PostingLine Line(string account, string dc, bool postLine) => new(isPosted, "CP", sj ? "MS" : "UHC", T(x, "BatchNo"), period,
                pv?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) ?? "", T(x, "PVCLNNo"), account, amount, dc, "4" + T(x, "ClinicCode"), T(x, "PayCode"),
                T(x, "ClinicCode"), postLine, pv, desc);
            if (kind is "G" or "C") { lines.Add(Line(sj ? "7520" : "7610", "C", false)); credit += amount; }
            if (kind is "G" or "D") { lines.Add(Line("84" + T(x, "ClinicCode"), "D", kind == "G")); debit += amount; }
        }
        return new PostingResult(lines, debit, credit, rows.Count, posted);
    }

    private static PostPayClnFilter Filter(PostPayClnRequest r, string? batch)
    {
        var clinic = (r.Clinic ?? "").Trim();
        var i = clinic.IndexOf('-');
        return new PostPayClnFilter(batch, clinic.Length == 0 ? null : (i > 0 ? clinic[..i] : clinic).Trim(), Date(r.PvDateFrom), Date(r.PvDateTo), Blank(r.PvFrom),
            Blank(r.PvTo), ((r.Posted ?? "0").Trim() + "0")[..1], Date(r.PostDateFrom), Date(r.PostDateTo), Blank(r.BatchFrom), Blank(r.BatchTo), r.Entity == 1);
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
