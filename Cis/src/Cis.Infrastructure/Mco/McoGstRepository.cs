using Cis.Application.Abstractions;
using Cis.Application.Mco;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Mco;

/// <summary><see cref="IMcoGstRepository"/> on the cis.McoGst_* procedures.</summary>
public sealed class McoGstRepository : IMcoGstRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> ListAsync(IDbSession session, McoGstFilter f) =>
        session.RowsAsync("cis.McoGst_List", new
        {
            mode = f.Mode, bordxDate = f.BordxDate, batchNo = f.BatchNo, from = f.From, to = f.To, invoiceNo = f.InvoiceNo, groupCode = f.GroupCode,
            subsidiary = f.Subsidiary, payor = f.Payor,
        }, 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SummaryAsync(IDbSession session, McoGstFilter f) => session.RowsAsync("cis.McoGst_Summary", P(f), 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> DetailsAsync(IDbSession session, McoGstFilter f) => session.RowsAsync("cis.McoGst_Details", P(f), 300);

    /// <inheritdoc />
    public async Task<List<string>> SubsidiariesAsync(IDbSession session, string company) =>
        (await session.RowsAsync("cis.McoGst_Subsidiaries", new { company })).Select(r => Convert.ToString(r["Subsidiary"])!).ToList();

    private static object P(McoGstFilter f) =>
        new { mode = f.Mode, bordxDate = f.BordxDate, batchNo = f.BatchNo, from = f.From, to = f.To, invoiceNo = f.InvoiceNo, payor = f.Payor };
}
