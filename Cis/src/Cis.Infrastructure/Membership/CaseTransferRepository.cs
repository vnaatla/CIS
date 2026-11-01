using Cis.Application.Abstractions;
using Cis.Application.Membership;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Membership;

/// <summary><see cref="ICaseTransferRepository"/> on the <c>cis.CaseTransfer_*</c> procedures.</summary>
public sealed class CaseTransferRepository : ICaseTransferRepository
{
    /// <inheritdoc />
    public async Task<TransferCase?> CaseAsync(IDbSession session, string casNumber)
    {
        var r = (await session.RowsAsync("cis.CaseTransfer_Case", new { casNumber })).FirstOrDefault();
        if (r is null) return null;
        var cov = r["CoverId"] as string ?? "";
        var patient = cov == "00" ? r["MbmName"] as string : r["MbmcName"] as string;
        return new TransferCase((string)r["CASNumber"]!, r["MbmNumber"] as string ?? "", cov, patient ?? "", r["DateVisit"] as DateTime?);
    }

    /// <inheritdoc />
    public async Task<string?> MemberNameAsync(IDbSession session, string mbmNumber, string covId) =>
        (await session.RowsAsync("cis.CaseTransfer_MemberName", new { mbmNumber, covId })).FirstOrDefault()?["Name"] as string;

    /// <inheritdoc />
    public Task<int> SaveAsync(IDbSession session, string casNumber, string mbmNumber, string covId) =>
        session.ScalarAsync<int>("cis.CaseTransfer_Save", new { casNumber, mbmNumber, covId });
}
