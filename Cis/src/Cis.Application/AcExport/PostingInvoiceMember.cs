using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.AcExport;

/// <summary>Invoice Posting > Member search (frmPostingInvoiceMember SearchRecord).</summary>
public sealed record PostInvMbmRequest
{
#pragma warning disable CS1591
    /// <summary>G: all lines, D: debit only, C: credit only.</summary>
    public string? Kind { get; init; } = "G";
    public string? GroupCompany { get; init; }
    public string? InvDateFrom { get; init; }
    public string? InvDateTo { get; init; }
    public string? CaseFrom { get; init; }
    public string? CaseTo { get; init; }
    /// <summary>"0 - No" / "1 - Yes".</summary>
    public string? Posted { get; init; } = "0 - No";
    public string? PostDateFrom { get; init; }
    public string? PostDateTo { get; init; }
    public string? BatchFrom { get; init; }
    public string? BatchTo { get; init; }
    public string? BordxDateFrom { get; init; }
    public string? BordxDateTo { get; init; }
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    /// <summary>The invoice registration period ("MM/YYYY").</summary>
    public string? RegPeriod { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PostInvMbmRequestValidator : Common.RequestValidator<PostInvMbmRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostInvMbmRequestValidator()
    {
        RuleFor(r => r.GroupCompany).MaximumLength(200);
        RuleFor(r => r.CaseFrom).MaximumLength(30);
        RuleFor(r => r.CaseTo).MaximumLength(30);
        RuleFor(r => r.BatchFrom).MaximumLength(20);
        RuleFor(r => r.BatchTo).MaximumLength(20);
        RuleFor(r => r.BordxFrom).MaximumLength(20);
        RuleFor(r => r.BordxTo).MaximumLength(20);
        RuleFor(r => r.RegPeriod).MaximumLength(7);
    }
}

/// <summary>A ticked case (and the batch no. its row listed, which the desktop writes as InvoiceNo).</summary>
public sealed record PostInvMbmTick(string CasNumber, string? BatchNo);

/// <summary>Post: the search, the accounting period and the ticked cases.</summary>
public sealed record PostInvMbmPostRequest(PostInvMbmRequest Search, string? Period, List<PostInvMbmTick> Ticked);

/// <summary>Lengths.</summary>
public sealed class PostInvMbmPostRequestValidator : Common.RequestValidator<PostInvMbmPostRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostInvMbmPostRequestValidator()
    {
        RuleFor(r => r.Period).MaximumLength(7);
        RuleFor(r => r.Search).SetValidator(new PostInvMbmRequestValidator());
        RuleFor(r => r.Ticked).Must(t => t.Count <= 20000 && t.All(x => x.CasNumber.Length <= 30 && (x.BatchNo ?? "").Length <= 20));
    }
}

/// <summary>The parsed criteria.</summary>
public sealed record PostInvMbmFilter(string? BatchNo, string? Company, DateTime? InvDateFrom, DateTime? InvDateTo, string? CaseFrom, string? CaseTo, string Posted,
    DateTime? PostDateFrom, DateTime? PostDateTo, string? BatchFrom, string? BatchTo, DateTime? BordxDateFrom, DateTime? BordxDateTo, string? BordxFrom, string? BordxTo,
    string? RegMonth, string? RegYear);

/// <summary>Data access.</summary>
public interface IPostInvMbmRepository
{
#pragma warning disable CS1591
    Task<List<string>> CompaniesAsync(IDbSession session);
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostInvMbmFilter filter);
    Task InsertAsync(IDbSession session, PostInvMbmTick tick, string accMonth, string accYear, string batchNo, string user);
#pragma warning restore CS1591
}

/// <summary>A/C Export > Invoice Posting > Member (frmPostingInvoiceMember).</summary>
public interface IPostInvMbmService
{
    /// <summary>The group companies.</summary>
    Task<List<string>> CompaniesAsync();

    /// <summary>Search / Search Debit / Search Credit.</summary>
    Task<PostingResult> SearchAsync(PostInvMbmRequest request);

    /// <summary>CmdPost_Click (after the confirmation): the batch no., its lines and the posting workbook.</summary>
    Task<(PostingResult Result, FileOutput File)> PostAsync(PostInvMbmPostRequest request, string user);
}

/// <inheritdoc />
public sealed class PostInvMbmService(IDbSessions sessions, IPostInvMbmRepository repository, IPostInvClnGstRepository numbers, IPostInvClnService export)
    : IPostInvMbmService
{
    /// <summary>Desktop text (the others are those of <see cref="PostInvClnService"/>).</summary>
    public const string NoPeriod = "Please Keyin for Accounting Period ! ";

    /// <inheritdoc />
    public async Task<List<string>> CompaniesAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.CompaniesAsync(s);
    }

    /// <inheritdoc />
    public async Task<PostingResult> SearchAsync(PostInvMbmRequest r)
    {
        await using var s = await sessions.OpenAsync();
        var result = Lines(await repository.SearchAsync(s, Filter(r, null)), Kind(r.Kind));
        if (result.Records == 0) throw new NotFoundException(PostInvClnService.NoRecord);
        return result;
    }

    /// <inheritdoc />
    public async Task<(PostingResult Result, FileOutput File)> PostAsync(PostInvMbmPostRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(PostInvClnService.NoSelection);
        var bordxFrom = Date(r.Search.BordxDateFrom);
        var bordxTo = Date(r.Search.BordxDateTo);
        if (bordxFrom is null || bordxTo is null) throw new BusinessRuleException(PostInvClnService.NoBordxDate);
        var period = (r.Period ?? "").Trim();
        if (period.Length < 7) throw new BusinessRuleException(NoPeriod);
        await using var tx = await sessions.BeginAsync();
        var open = (await repository.SearchAsync(tx, Filter(r.Search with { Posted = "0 - No" }, null))).Select(x => T(x, "CASNumber")).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (r.Ticked.Any(t => !open.Contains(t.CasNumber.Trim()))) throw new BusinessRuleException(PostInvClnService.Posted);
        var batch = await numbers.NextBatchAsync(tx);
        foreach (var t in r.Ticked.DistinctBy(t => t.CasNumber.Trim()))
            await repository.InsertAsync(tx, t with { CasNumber = t.CasNumber.Trim() }, period[..2], period.Substring(3, 4), batch, user);
        var result = Lines(await repository.SearchAsync(tx, new PostInvMbmFilter(batch, null, null, null, null, null, "", null, null, null, null, null, null, null, null, null,
            null)), "G") with { BatchNo = batch };
        await tx.CommitAsync();
        return (result, export.Export(result, $"{batch} RC-MBMInv {bordxFrom:yyyyMMdd}-{bordxTo:dd}"));
    }

    /// <summary>MBMInvListing: debit 6000 (actual), credit 82 + group company (actual − postage) and, with postage, credit 4040 (postage); the last line is the one Post
    /// writes. Search Debit / Credit list one side of the same lines (the desktop's Search Credit listed the 82 line twice, with the actual and with actual − postage,
    /// when there was postage).</summary>
    public static PostingResult Lines(List<Dictionary<string, object?>> rows, string kind)
    {
        var lines = new List<PostingLine>();
        decimal debit = 0, credit = 0;
        var posted = false;
        foreach (var x in rows)
        {
            var isPosted = T(x, "InvMBMStatus") == "1";
            posted |= isPosted;
            var actual = N(x, "CLMTotalActual");
            var postage = N(x, "CLMPostageFee");
            var member = T(x, "MBMNumber");
            var star = member.IndexOf('*');
            var last4 = (star < 0 ? member : member[..star]) is var m && m.Length > 4 ? m[^4..] : (star < 0 ? member : member[..star]);
            var desc = T(x, "PatientName");
            if (desc.Length > 25) desc = desc[..25];
            var period = T(x, "AccMonth") + T(x, "AccYear") == "" ? "" : $"{T(x, "AccMonth")}/{T(x, "AccYear")}";
            var inv = x.GetValueOrDefault("InvoiceDate") as DateTime?;
            var bordxDate = (x.GetValueOrDefault("BordxDate") as DateTime?)?.ToString("yyMM", CultureInfo.InvariantCulture) ?? "";
            var set = new List<(string Account, string Dc, decimal Amount)>();
            if (kind is "G" or "D") set.Add(("6000", "D", actual));
            if (kind is "G" or "C")
            {
                set.Add(("82" + T(x, "GRPCoCode"), "C", actual - postage));
                if (postage > 0) set.Add(("4040", "C", postage));
            }
            for (var aa = 1; aa <= set.Count; aa++)
            {
                var (account, dc, amount) = set[aa - 1];
                lines.Add(new PostingLine(isPosted, "RC", "UHC", T(x, "BatchNo"), period, inv?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) ?? "", T(x, "CASNumber"),
                    account, amount, dc, "1R", T(x, "PAYCode"), T(x, "GRPCoCode"), kind == "G" && aa == set.Count, inv, desc)
                {
                    Extra = new Dictionary<string, string> { ["BordxPeriod"] = bordxDate, ["Bordx"] = T(x, "INVBordxRefNo"), ["Member"] = last4 },
                });
                if (dc == "C") credit += amount; else debit += amount;
            }
        }
        return new PostingResult(lines, debit, credit, rows.Count, posted);
    }

    private static PostInvMbmFilter Filter(PostInvMbmRequest r, string? batch)
    {
        var reg = (r.RegPeriod ?? "").Trim();
        return new PostInvMbmFilter(batch, Blank(r.GroupCompany), Date(r.InvDateFrom), Date(r.InvDateTo), Blank(r.CaseFrom), Blank(r.CaseTo),
            ((r.Posted ?? "0").Trim() + "0")[..1], Date(r.PostDateFrom), Date(r.PostDateTo), Blank(r.BatchFrom), Blank(r.BatchTo), Date(r.BordxDateFrom), Date(r.BordxDateTo),
            Blank(r.BordxFrom), Blank(r.BordxTo), reg.Length >= 7 ? reg[..2] : null, reg.Length >= 7 ? reg.Substring(3, 4) : null);
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
