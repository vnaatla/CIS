using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;

namespace Cis.Application.AcExport;

/// <summary>The parsed criteria of Invoice Posting > Clinic (GST).</summary>
public sealed record PostInvClnGstFilter(string? BatchNo, string? Clinic, DateTime? InvDateFrom, DateTime? InvDateTo, string? InvoiceFrom, string? InvoiceTo, string Posted,
    DateTime? PostDateFrom, DateTime? PostDateTo, string? BatchFrom, string? BatchTo, string? SunFrom, string? SunTo, string? AccMonth, string? AccYear);

/// <summary>Data access.</summary>
public interface IPostInvClnGstRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PostInvClnGstFilter filter);
    Task<string> NextBatchAsync(IDbSession session);
    Task InsertAsync(IDbSession session, PostInvClnTick tick, string accMonth, string accYear, string invBatch, string batchNo, string user);
#pragma warning restore CS1591
}

/// <summary>A/C Export > Invoice Posting > Clinic (GST) (frmPostingInvoiceClnGST).</summary>
public interface IPostInvClnGstService
{
    /// <summary>Search / Search Debit / Search Credit.</summary>
    Task<PostingResult> SearchAsync(PostInvClnRequest request);

    /// <summary>CmdPost_Click (after the confirmation): the batch no., its lines and the posting workbook.</summary>
    Task<(PostingResult Result, FileOutput File)> PostAsync(PostInvClnPostRequest request, string user);
}

/// <inheritdoc />
public sealed class PostInvClnGstService(IDbSessions sessions, IPostInvClnGstRepository repository, IPostInvClnRepository numbers, IPostInvClnService export,
    TimeProvider clock) : IPostInvClnGstService
{
    /// <summary>Desktop text (the others are those of <see cref="PostInvClnService"/>).</summary>
    public const string NoInvDate = "Please Enter Invoice Date! ";

    /// <inheritdoc />
    public async Task<PostingResult> SearchAsync(PostInvClnRequest r)
    {
        await using var s = await sessions.OpenAsync();
        var result = Lines(await repository.SearchAsync(s, Filter(r, null)), Kind(r.Kind), Today);
        if (result.Records == 0) throw new NotFoundException(PostInvClnService.NoRecord);
        return result;
    }

    private DateTime Today => clock.GetLocalNow().Date;

    /// <inheritdoc />
    public async Task<(PostingResult Result, FileOutput File)> PostAsync(PostInvClnPostRequest r, string user)
    {
        if (r.Ticked.Count == 0) throw new BusinessRuleException(PostInvClnService.NoSelection);
        var invFrom = Date(r.Search.InvDateFrom);
        var invTo = Date(r.Search.InvDateTo);
        if (invFrom is null || invTo is null) throw new BusinessRuleException(NoInvDate);
        var period = (r.Period ?? "").Trim();
        if (period.Length < 7) throw new BusinessRuleException(PostInvClnService.NoPeriod);
        await using var tx = await sessions.BeginAsync();
        var open = (await repository.SearchAsync(tx, Filter(r.Search with { Posted = "0 - No" }, null)))
            .ToDictionary(x => (T(x, "InvoiceNo"), T(x, "ClinicCode")), x => x);
        if (r.Ticked.Any(t => !open.ContainsKey((t.InvoiceNo.Trim(), t.ClinicCode.Trim())))) throw new BusinessRuleException(PostInvClnService.Posted);
        var batch = await repository.NextBatchAsync(tx);
        foreach (var t in r.Ticked.DistinctBy(t => (t.InvoiceNo.Trim(), t.ClinicCode.Trim())))
        {
            var row = open[(t.InvoiceNo.Trim(), t.ClinicCode.Trim())];
            // An invoice with an FFS fee split into GST gets a "T" SUN invoice batch, the others "I".
            var invBatch = (Split(N(row, "CLMMCOProcFee"), Today).NoGst > 0 ? "T" : "I") + await numbers.NextInvNoAsync(tx);
            await repository.InsertAsync(tx, t with { InvoiceNo = t.InvoiceNo.Trim(), ClinicCode = t.ClinicCode.Trim() }, period[..2], period.Substring(3, 4), invBatch, batch,
                user);
        }
        var result = Lines(await repository.SearchAsync(tx, new PostInvClnGstFilter(batch, null, null, null, null, null, "", null, null, null, null, null, null, null, null)),
            "G", Today) with { BatchNo = batch };
        await tx.CommitAsync();
        return (result, export.Export(result, $"{batch} RC-CLNInv {invFrom:yyyyMMdd}-{invTo:dd}"));
    }

    /// <summary>The FFS fee split: from 01/04/2015, 94.34 % without GST and the rest GST.</summary>
    public static (decimal NoGst, decimal Gst) Split(decimal ffs, DateTime today)
    {
        if (ffs == 0 || today < new DateTime(2015, 4, 1)) return (0, 0);
        var noGst = Math.Round(ffs * 0.9434m, 2, MidpointRounding.AwayFromZero);
        return (noGst, Math.Round(ffs - noGst, 2, MidpointRounding.AwayFromZero));
    }

    /// <summary>ClnInvListing (GST): no FFS fee → credit 84 + clinic and debit 6000 (payable); with an FFS fee → debit 6000 (actual − co-pay), credit 84 + clinic
    /// (payable), credit 1002 (the FFS fee without GST, or the whole fee before GST) and credit 8830 (the GST). The totals add the payable amount for every line,
    /// as the desktop.</summary>
    public static PostingResult Lines(List<Dictionary<string, object?>> rows, string kind, DateTime today)
    {
        var lines = new List<PostingLine>();
        decimal debit = 0, credit = 0;
        var posted = false;
        var date = today.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        static decimal R(decimal v) => Math.Round(v, 2, MidpointRounding.AwayFromZero);
        foreach (var x in rows)
        {
            var isPosted = T(x, "InvCLNStatus") == "1";
            posted |= isPosted;
            var payable = R(N(x, "CLMPayableByMedix2C"));
            var actual = R(N(x, "CLMTotalActual") - R(N(x, "CLMCoPayColl")));
            var ffs = N(x, "CLMMCOProcFee");
            var (noGst, gst) = Split(ffs, today);
            var period = T(x, "AccMonth") + T(x, "AccYear") == "" ? "" : $"{T(x, "AccMonth")}/{T(x, "AccYear")}";
            var inv = x.GetValueOrDefault("InvoiceDate") as DateTime?;
            PostingLine Line(string account, string dc, decimal amount, bool postLine) => new(isPosted, "RC", "UHC", T(x, "CLNInvoiceBatch"), period, date,
                T(x, "InvoiceNo"), account, amount, dc, "4" + T(x, "ClinicCode"), T(x, "PAYCode"), T(x, "ClinicCode"), postLine, inv, ffs == 0 ? T(x, "ClnName") : "FFS");
            var set = ffs == 0
                ? new List<PostingLine> { Line("84" + T(x, "ClinicCode"), "C", payable, false), Line("6000", "D", payable, true) }
                : new List<PostingLine> { Line("6000", "D", actual, true), Line("84" + T(x, "ClinicCode"), "C", payable, false), Line("1002", "C", noGst != 0 ? noGst : R(ffs), false) };
            if (ffs != 0 && noGst != 0) set.Add(Line("8830", "C", gst, false));
            foreach (var l in set.Where(l => kind == "G" || l.DebitCredit == kind))
            {
                lines.Add(kind == "G" ? l : l with { PostLine = false });
                if (l.DebitCredit == "C") credit += payable; else debit += payable;
            }
        }
        return new PostingResult(lines, debit, credit, rows.Count, posted);
    }

    private static PostInvClnGstFilter Filter(PostInvClnRequest r, string? batch)
    {
        var clinic = (r.Clinic ?? "").Trim();
        var i = clinic.IndexOf('-');
        var period = (r.InvoicePeriod ?? "").Trim();
        return new PostInvClnGstFilter(batch, clinic.Length == 0 ? null : (i > 0 ? clinic[..i] : clinic).Trim(), Date(r.InvDateFrom), Date(r.InvDateTo), Blank(r.InvoiceFrom),
            Blank(r.InvoiceTo), ((r.Posted ?? "0").Trim() + "0")[..1], Date(r.PostDateFrom), Date(r.PostDateTo), Blank(r.BatchFrom), Blank(r.BatchTo), Blank(r.SunFrom),
            Blank(r.SunTo), period.Length >= 7 ? period[..2] : null, period.Length >= 7 ? period.Substring(3, 4) : null);
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
