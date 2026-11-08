using Cis.Application.Abstractions;
using Cis.Application.Mco;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Mco;

/// <summary><see cref="IMcoZeroRepository"/> on cis.McoZero_List / cis.McoZero_Company.</summary>
public sealed class McoZeroRepository : IMcoZeroRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> ListAsync(IDbSession session, McoGstFilter f, string? policyNo) =>
        session.RowsAsync("cis.McoZero_List", new
        {
            mode = f.Mode, bordxDate = f.BordxDate, batchNo = f.BatchNo, from = f.From, to = f.To, invoiceNo = f.InvoiceNo, groupCode = f.GroupCode,
            subsidiary = f.Subsidiary, payor = f.Payor, policyNo,
        }, 300);

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> CompanyAsync(IDbSession session, string? subsidiary, string? subsidiary2, string? group) =>
        (await session.RowsAsync("cis.McoZero_Company", new { subsidiary, subsidiary2, group })).FirstOrDefault();
}
