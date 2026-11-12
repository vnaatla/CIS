using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.AcExport;

/// <summary>Payment Posting > Member search (FrmPostingPay2MBM SearchRecord).</summary>
public sealed record PostPayMbmRequest
{
#pragma warning disable CS1591
    /// <summary>G: all lines, D: debit only, C: credit only.</summary>
    public string? Kind { get; init; } = "G";
    public string? GroupCompany { get; init; }
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
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PostPayMbmRequestValidator : Common.RequestValidator<PostPayMbmRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostPayMbmRequestValidator()
    {
        RuleFor(r => r.GroupCompany).MaximumLength(200);
        RuleFor(r => r.PvFrom).MaximumLength(30);
        RuleFor(r => r.PvTo).MaximumLength(30);
        RuleFor(r => r.BatchFrom).MaximumLength(20);
        RuleFor(r => r.BatchTo).MaximumLength(20);
    }
}

/// <summary>Post: the search, the accounting period and the ticked PV nos.</summary>
public sealed record PostPayMbmPostRequest(PostPayMbmRequest Search, string? Period, List<string> Ticked);

/// <summary>Lengths.</summary>
public sealed class PostPayMbmPostRequestValidator : Common.RequestValidator<PostPayMbmPostRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostPayMbmPostRequestValidator()
    {
        RuleFor(r => r.Period).MaximumLength(7);
        RuleFor(r => r.Search).SetValidator(new PostPayMbmRequestValidator());
        RuleFor(r => r.Ticked).Must(t => t.Count <= 20000 && t.All(x => x.Length <= 30));
    }
}

/// <summary>The parsed criteria.</summary>
public sealed record PostPayMbmFilter(string? BatchNo, string? Company, DateTime? PvDateFrom, DateTime? PvDateTo, string? PvFrom, string? PvTo, string Posted,
    DateTime? PostDateFrom, DateTime? PostDateTo, string? BatchFrom, string? BatchTo);

/// <summary>Data access.</summary>
public interface IPostPayMbmRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostPayMbmFilter filter);
    Task InsertAsync(IDbSession session, string transRef, string accMonth, string accYear, string batchNo, string user);
#pragma warning restore CS1591
}

/// <summary>A/C Export > Payment Posting > Member (FrmPostingPay2MBM).</summary>
public interface IPostPayMbmService
{
    /// <summary>Search / Search Debit / Search Credit.</summary>
    Task<PostingResult> SearchAsync(PostPayMbmRequest request);

    /// <summary>CmdPost_Click (after the confirmation): the batch no., its lines and the posting workbook.</summary>
    Task<(PostingResult Result, FileOutput File)> PostAsync(PostPayMbmPostRequest request, string user);
}

/// <inheritdoc />
public sealed class PostPayMbmService(IDbSessions sessions, IPostPayMbmRepository repository, IPostInvClnGstRepository numbers, IPostInvClnService export)
    : IPostPayMbmService
{
    /// <inheritdoc />
    public async Task<PostingResult> SearchAsync(PostPayMbmRequest r)
    {
        await using var s = await sessions.OpenAsync();
        var result = Lines(await repository.SearchAsync(s, Filter(r, null)), Kind(r.Kind));
        if (result.Records == 0) throw new NotFoundException(PostInvClnService.NoRecord);
        return result;
    }

    /// <inheritdoc />
    public async Task<(PostingResult Result, FileOutput File)> PostAsync(PostPayMbmPostRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(PostInvClnService.NoSelection);
        var pvFrom = Date(r.Search.PvDateFrom);
        var pvTo = Date(r.Search.PvDateTo);
        if (pvFrom is null || pvTo is null) throw new BusinessRuleException(PostPayClnService.NoPvDate);
        var period = (r.Period ?? "").Trim();
        if (period.Length < 7) throw new BusinessRuleException(PostPayClnService.NoPeriod);
        await using var tx = await sessions.BeginAsync();
        var open = (await repository.SearchAsync(tx, Filter(r.Search with { Posted = "0 - No" }, null))).Select(x => T(x, "PVNo"))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (r.Ticked.Any(t => !open.Contains(t.Trim()))) throw new BusinessRuleException(PostInvClnService.Posted);
        var batch = await numbers.NextBatchAsync(tx);
        foreach (var t in r.Ticked.Select(t => t.Trim()).Distinct())
            await repository.InsertAsync(tx, t, period[..2], period.Substring(3, 4), batch, user);
        var result = Lines(await repository.SearchAsync(tx, new PostPayMbmFilter(batch, null, null, null, null, null, "", null, null, null, null)), "G") with { BatchNo = batch };
        await tx.CommitAsync();
        return (result, export.Export(result, $"{batch} CP-PayMBM {pvFrom:yyyyMMdd}-{pvTo:dd}"));
    }

    /// <summary>MBMPayListing: per PV a credit line 7610 (the payee; the line Post writes) and a debit line 82 + group company; journal CP, analysis 1R. Search Debit
    /// lists the 82 line, Search Credit the 7610 line.</summary>
    public static PostingResult Lines(List<Dictionary<string, object?>> rows, string kind)
    {
        var lines = new List<PostingLine>();
        decimal debit = 0, credit = 0;
        var posted = false;
        foreach (var x in rows)
        {
            var isPosted = T(x, "Pay2MBMStatus") == "1";
            posted |= isPosted;
            var amount = N(x, "MBMPaidAmt");
            var pv = x.GetValueOrDefault("PVDate") as DateTime?;
            var period = T(x, "AccMonth") + T(x, "AccYear") == "" ? "" : $"{T(x, "AccMonth")}/{T(x, "AccYear")}";
            var desc = T(x, "PVPayto");
            if (desc.Length > 25) desc = desc[..25];
            PostingLine Line(string account, string dc, bool postLine) => new(isPosted, "CP", "UHC", T(x, "BatchNo"), period,
                pv?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) ?? "", T(x, "PVNo"), account, amount, dc, "1R", T(x, "PAYCode"), T(x, "GRPCoCode"), postLine,
                pv, desc);
            if (kind == "G")
            {
                lines.Add(Line("7610", "C", true));
                lines.Add(Line("82" + T(x, "GRPCoCode"), "D", false));
                credit += amount;
                debit += amount;
            }
            else
            {
                lines.Add(kind == "D" ? Line("82" + T(x, "GRPCoCode"), "D", false) : Line("7610", "C", false));
                // The desktop added Search Debit's 82 lines to the credit total.
                if (kind == "D") debit += amount; else credit += amount;
            }
        }
        return new PostingResult(lines, debit, credit, rows.Count, posted);
    }

    private static PostPayMbmFilter Filter(PostPayMbmRequest r, string? batch) =>
        new(batch, Blank(r.GroupCompany), Date(r.PvDateFrom), Date(r.PvDateTo), Blank(r.PvFrom), Blank(r.PvTo), ((r.Posted ?? "0").Trim() + "0")[..1], Date(r.PostDateFrom),
            Date(r.PostDateTo), Blank(r.BatchFrom), Blank(r.BatchTo));

    private static string Kind(string? k) => (k ?? "G").Trim().ToUpperInvariant() switch { "D" => "D", "C" => "C", _ => "G" };

    private static DateTime? Date(string? v) =>
        string.IsNullOrWhiteSpace(v) ? null : ClinicInvoiceRules.Date(v) ?? throw new BusinessRuleException(PostInvClnService.BadDate);

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
