using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IBatchTransferRepository"/> on the <c>cis.BatchTransfer_*</c> procedures.</summary>
public sealed class BatchTransferRepository : IBatchTransferRepository
{
    /// <inheritdoc />
    public Task<List<BatchRow>> BatchAsync(IDbSession session, string batchNo) =>
        session.QueryAsync<BatchRow>("cis.BatchTransfer_Batch", new { batchNo });

    /// <inheritdoc />
    public Task<int> SaveAsync(IDbSession session, string fromBatch, string toBatch, string remark, decimal amount) =>
        session.ScalarAsync<int>("cis.BatchTransfer_Save", new { fromBatch, toBatch, remark, amount });
}
