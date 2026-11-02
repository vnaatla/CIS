using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IOpBatchRepository"/> on the <c>cis.OpBatch_*</c> procedures.</summary>
public sealed class OpBatchRepository : IOpBatchRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, OpBatchFilters f)
    {
        static string? T(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        static string? First(string? v) => T(v) is { } t ? t[..1].ToUpperInvariant() : null;
        static decimal? Amount(string? v) => decimal.TryParse(v?.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : null;
        var clinic = T(f.Clinic);
        if (clinic is not null && clinic.IndexOf('-') is var dash and > 0) clinic = clinic[..dash].Trim();
        var payTo = T(f.PayTo);
        if (payTo is not null) payTo = (payTo.IndexOf('-') is var p and > 0 ? payTo[..p] : payTo.Length > 4 ? payTo[..4] : payTo).Trim();
        var referral = T(f.Referral)?.ToUpperInvariant() switch { "YES" => "Y", "NO" => "N", _ => null };
        return session.RowsAsync("cis.OpBatch_Search", new
        {
            noBatch = f.NoBatch, noBordx = f.NoBordx, bordxFrom = T(f.BordxFrom), bordxTo = T(f.BordxTo), noBtbgBordx = f.NoBtbgBordx,
            btbgFrom = T(f.BtbgFrom), btbgTo = T(f.BtbgTo), casFrom = T(f.CasFrom), casTo = T(f.CasTo), clinicCode = clinic, groupCompany = T(f.GroupCompany),
            invoiceNo = T(f.InvoiceNo), visitFrom = OpBatchService.Date(f.VisitFrom), visitTo = OpBatchService.Date(f.VisitTo), noInvoice = f.NoInvoice,
            invoiceFrom = OpBatchService.Date(f.InvoiceFrom), invoiceTo = OpBatchService.Date(f.InvoiceTo), amountFrom = Amount(f.AmountFrom),
            amountTo = Amount(f.AmountTo), batchFrom = T(f.BatchFrom), batchTo = T(f.BatchTo), omni = f.Omni, regFrom = OpBatchService.Date(f.RegFrom),
            regTo = OpBatchService.Date(f.RegTo), clinicType = First(f.ClinicType), ffs = First(f.Ffs), hasBatch = First(f.HasBatch), payTo, referral,
            entity = f.Entity,
        });
    }

    /// <inheritdoc />
    public Task<bool> InvoiceUsedAsync(IDbSession session, string invoiceNo) => session.ScalarAsync<bool>("cis.OpBatch_InvoiceUsed", new { invoiceNo });

    /// <inheritdoc />
    public async Task<string> NextBatchAsync(IDbSession session, string opCode) =>
        (await session.ScalarAsync<string>("cis.OpBatch_Next", new { opCode }))!;

    /// <inheritdoc />
    public Task SaveAsync(IDbSession session, string casNumber, string? batchNo, string userCode, DateTime today, string? invoiceNo, DateTime? invoiceDate) =>
        session.ScalarAsync<int>("cis.OpBatch_Save", new { casNumber, batchNo, userCode, today, invoiceNo, invoiceDate });
}
