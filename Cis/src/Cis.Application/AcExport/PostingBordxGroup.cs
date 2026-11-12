using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.AcExport;

/// <summary>Bordx Posting > Grp Company search (frmPostingBordxtoGrpComp / frmPostingBordxtoGrpGST SearchRecord).</summary>
public sealed record PostBordxGrpRequest
{
#pragma warning disable CS1591
    /// <summary>G: all lines, D: debit only, C: credit only.</summary>
    public string? Kind { get; init; } = "G";
    public string? GroupCompany { get; init; }
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    public string? BordxDateFrom { get; init; }
    public string? BordxDateTo { get; init; }
    public string? PostDateFrom { get; init; }
    public string? PostDateTo { get; init; }
    /// <summary>"0 - No" / "1 - Yes".</summary>
    public string? Posted { get; init; } = "0 - No";
    public string? BatchFrom { get; init; }
    public string? BatchTo { get; init; }
    /// <summary>GST screen: 0 Health Connect (payors other than SJ), 1 Sihat Johor (SJ).</summary>
    public int Entity { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PostBordxGrpRequestValidator : Common.RequestValidator<PostBordxGrpRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostBordxGrpRequestValidator()
    {
        RuleFor(r => r.GroupCompany).MaximumLength(200);
        RuleFor(r => r.BordxFrom).MaximumLength(20);
        RuleFor(r => r.BordxTo).MaximumLength(20);
        RuleFor(r => r.BatchFrom).MaximumLength(20);
        RuleFor(r => r.BatchTo).MaximumLength(20);
    }
}

/// <summary>Post: the search, the accounting period and the ticked bordx nos.</summary>
public sealed record PostBordxGrpPostRequest(PostBordxGrpRequest Search, string? Period, List<string> Ticked);

/// <summary>Lengths.</summary>
public sealed class PostBordxGrpPostRequestValidator : Common.RequestValidator<PostBordxGrpPostRequest>
{
    /// <summary>Defines the rules.</summary>
    public PostBordxGrpPostRequestValidator()
    {
        RuleFor(r => r.Period).MaximumLength(7);
        RuleFor(r => r.Search).SetValidator(new PostBordxGrpRequestValidator());
        RuleFor(r => r.Ticked).Must(t => t.Count <= 20000 && t.All(x => x.Length <= 20));
    }
}

/// <summary>The parsed criteria.</summary>
public sealed record PostBordxGrpFilter(string? BatchNo, string? Company, string? BordxFrom, string? BordxTo, DateTime? BordxDateFrom, DateTime? BordxDateTo,
    DateTime? PostDateFrom, DateTime? PostDateTo, string Posted, string? BatchFrom, string? BatchTo, bool? Sj);

/// <summary>Data access.</summary>
public interface IPostBordxGrpRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostBordxGrpFilter filter);
    Task InsertAsync(IDbSession session, string bordxNo, string accMonth, string accYear, string batchNo, string user);
#pragma warning restore CS1591
}

/// <summary>A/C Export > Bordx Posting > Grp Company (frmPostingBordxtoGrpComp, hidden) and Grp Company (GST) (frmPostingBordxtoGrpGST, <c>gst</c>).</summary>
public interface IPostBordxGrpService
{
    /// <summary>Search / Search Debit / Search Credit.</summary>
    Task<PostingResult> SearchAsync(PostBordxGrpRequest request, bool gst);

    /// <summary>CmdExp_Click (after the confirmation): the batch no., its lines and the posting workbook.</summary>
    Task<(PostingResult Result, FileOutput File)> PostAsync(PostBordxGrpPostRequest request, string user, bool gst);
}

/// <inheritdoc />
public sealed class PostBordxGrpService(IDbSessions sessions, IPostBordxGrpRepository repository, IPostInvClnGstRepository numbers, IPostInvClnService export)
    : IPostBordxGrpService
{
    /// <summary>Desktop text.</summary>
    public const string NoRecord = "Sorry, No Record Found! ";

    private static readonly string[] Months =
        ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];

    /// <inheritdoc />
    public async Task<PostingResult> SearchAsync(PostBordxGrpRequest r, bool gst)
    {
        await using var s = await sessions.OpenAsync();
        var result = Lines(await repository.SearchAsync(s, Filter(r, null, gst)), Kind(r.Kind), gst);
        if (result.Records == 0) throw new NotFoundException(NoRecord);
        return result;
    }

    /// <inheritdoc />
    public async Task<(PostingResult Result, FileOutput File)> PostAsync(PostBordxGrpPostRequest r, string user, bool gst)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(PostInvClnService.NoSelection);
        var from = Date(r.Search.BordxDateFrom);
        var to = Date(r.Search.BordxDateTo);
        if (from is null || to is null) throw new BusinessRuleException(PostInvClnService.NoBordxDate);
        var period = (r.Period ?? "").Trim();
        if (period.Length < 7) throw new BusinessRuleException(PostPayClnService.NoPeriod);
        await using var tx = await sessions.BeginAsync();
        var open = (await repository.SearchAsync(tx, Filter(r.Search with { Posted = "0 - No" }, null, gst))).Select(x => T(x, "BordxRefNo"))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (r.Ticked.Any(t => !open.Contains(t.Trim()))) throw new BusinessRuleException(PostInvClnService.Posted);
        var batch = await numbers.NextBatchAsync(tx);
        foreach (var t in r.Ticked.Select(t => t.Trim()).Distinct())
            await repository.InsertAsync(tx, t, period[..2], period.Substring(3, 4), batch, user);
        var result = Lines(await repository.SearchAsync(tx, new PostBordxGrpFilter(batch, null, null, null, null, null, null, null, "", null, null, null)), "G", gst)
            with { BatchNo = batch };
        await tx.CommitAsync();
        return (result, export.Export(result, $"{batch} BI-BordxGrp {from:yyyyMMdd}-{to:dd}"));
    }

    /// <summary>GrpCompListing: per bordereaux a debit line (65 + the company / payor account) and a credit line 6000 (the line Post writes); the hidden screen adds a
    /// 1002 credit for the MCO fee of fee-for-service companies and posts the payable to clinic on 6000 then. Search Debit / Credit list actual − co-pay.</summary>
    public static PostingResult Lines(List<Dictionary<string, object?>> rows, string kind, bool gst)
    {
        var lines = new List<PostingLine>();
        decimal total = 0;
        var posted = false;
        static decimal R(decimal v) => Math.Round(v, 2, MidpointRounding.AwayFromZero);
        foreach (var x in rows)
        {
            var isPosted = T(x, "BordxGrpCompStatus") == "1";
            posted |= isPosted;
            var bordx = T(x, "BordxRefNo");
            var date = x.GetValueOrDefault("BordxDate") as DateTime?;
            var group = T(x, "MBMGroupCompany");
            var printMonth = Array.IndexOf(Months, T(x, "BORDXPrintMonth").ToUpperInvariant()) is var m and >= 0 ? (m + 1).ToString("00", CultureInfo.InvariantCulture) : "";
            var printYear = T(x, "BORDXPrintYear");
            var gstDesc = date is null ? group : $"{date.Value.Year}/{date.Value.Month}-{group}";
            var desc = gst ? gstDesc : $"{(printYear.Length >= 4 ? printYear.Substring(2, 2) : "")}{printMonth}-{group}";
            var bordxAmt = R(N(x, "CLMBordxAmt"));
            var mco = Math.Round(N(x, "CLMMCOProcFee"), 3, MidpointRounding.AwayFromZero);
            var ffs = mco != 0 && T(x, "GRPMCOStatus") == "F";
            var payTo = T(x, "INVPayTo") == "C" ? T(x, "PAYCode") : "1RE";
            string account;
            if (bordx.StartsWith("KDC", StringComparison.Ordinal) || (gst && bordx.StartsWith("KCN", StringComparison.Ordinal))) account = "65" + T(x, "CoGRPCode");
            else if (T(x, "PAYAccCode").Length > 0) account = "65" + (gst ? T(x, "PAYAccCode") : Mid(T(x, "PAYAccCode"), 3, 4));
            else if (gst && group == "GAS MALAYSIA BERHAD") account = "65A002";
            else account = "65" + T(x, "PAYCode");
            PostingLine Line(string acc, string dc, decimal amount, string text, bool postLine) => new(isPosted, "BI", "UHC", T(x, "BatchNo"),
                T(x, "AccMonth") + T(x, "AccYear") == "" ? "" : $"{T(x, "AccMonth")}/{T(x, "AccYear")}", date?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) ?? "",
                bordx, acc, amount, dc, "", payTo, T(x, "GRPCoCode"), postLine, date, gst ? text : (text.Length > 25 ? text[..25] : text))
            {
                Extra = new Dictionary<string, string>
                {
                    ["BordxPeriod"] = date?.ToString("yyMM", CultureInfo.InvariantCulture) ?? "", ["Bordx"] = Mid(bordx, 4, 5),
                },
            };
            if (kind == "G")
            {
                var debitDesc = gst && date is not null && T(x, "BordxSubsidary").Length > 0 ? $"{date.Value.Year}/{date.Value.Month}-{T(x, "BordxSubsidary")}" : desc;
                lines.Add(Line(account, "D", bordxAmt, debitDesc, false));
                lines.Add(Line("6000", "C", !gst && ffs && T(x, "INVPayTo") == "C" ? R(N(x, "CLMPayableByMedix2C")) : bordxAmt, desc, true));
                if (!gst && ffs) lines.Add(Line("1002", "C", R(mco), desc, false));
                total += bordxAmt;
            }
            else
            {
                var amount = R(N(x, "CLMTotalActual") - N(x, "CLMCoPayColl"));
                var text = gst ? gstDesc : $"{T(x, "BORDXPrintMonth")}{printYear}-{group}";
                lines.Add(kind == "C" ? Line("6000", "C", amount, text, false) : Line(T(x, "CoGRPAcccode"), "D", amount, text, false));
                total += amount;
            }
        }
        return new PostingResult(lines, total, total, rows.Count, posted);
    }

    private static string Mid(string v, int start, int length) => v.Length < start ? "" : v.Substring(start - 1, Math.Min(length, v.Length - start + 1));

    /// <summary>SearchRecord. The GST screen swaps the bordx no. boxes (the "from" box is the upper bound), as the desktop.</summary>
    private static PostBordxGrpFilter Filter(PostBordxGrpRequest r, string? batch, bool gst)
    {
        var from = Date(r.BordxDateFrom);
        var to = Date(r.BordxDateTo);
        var postFrom = Date(r.PostDateFrom);
        var postTo = Date(r.PostDateTo);
        var bothBordx = from is not null && to is not null;
        var bothPost = postFrom is not null && postTo is not null;
        var (bordxFrom, bordxTo) = gst
            ? (Blank(r.BordxFrom) is null ? null : Blank(r.BordxTo), Blank(r.BordxTo) is null ? null : Blank(r.BordxFrom))
            : (Blank(r.BordxFrom), Blank(r.BordxTo));
        var posted = (r.Posted ?? "").Trim();
        return new PostBordxGrpFilter(batch, Blank(r.GroupCompany), bordxFrom, bordxTo, bothBordx ? from : null, bothBordx ? to : null, bothPost ? postFrom : null,
            bothPost ? postTo : null, posted.Length > 0 ? posted[..1] : "", Blank(r.BatchFrom), Blank(r.BatchTo), gst ? r.Entity == 1 : null);
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
