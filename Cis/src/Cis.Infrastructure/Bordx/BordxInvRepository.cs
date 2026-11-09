using Cis.Application.Abstractions;
using Cis.Application.Bordx;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Bordx;

/// <summary><see cref="IBordxInvRepository"/> on cis.BordxInv_*.</summary>
public sealed class BordxInvRepository : IBordxInvRepository
{
    /// <inheritdoc />
    public async Task<(List<(string Code, string Name, string? PayCode)> Companies, List<(string Code, string Name)> Payors)> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.BordxInv_Lookups");
        var companies = (await grid.ReadAsync<(string Code, string Name, string? PayCode)>()).ToList();
        return (companies, (await grid.ReadAsync<(string Code, string Name)>()).ToList());
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, DateTime? from, DateTime? to, string? payCode, string? grpCode) =>
        session.RowsAsync("cis.BordxInv_Search", new { regFrom = from, regTo = to, payCode, grpCode }, 300);

    /// <inheritdoc />
    public Task<List<(string CasNumber, decimal? CLMFinalApprove)>> CasesAsync(IDbSession session, BordxInvTick t) =>
        session.QueryAsync<(string CasNumber, decimal? CLMFinalApprove)>("cis.BordxInv_Cases",
            new { invoiceNo = t.InvoiceNo, invoiceDate = t.InvoiceDate, registerDate = t.RegisterDate, grpCode = t.GrpCoCode });

    /// <inheritdoc />
    public Task UpdateCaseAsync(IDbSession session, string casNumber, string bordxNo) =>
        session.ExecAsync("cis.BordxInv_UpdateCase", new { casNumber, bordxNo });

    /// <inheritdoc />
    public Task SaveBordxAsync(IDbSession session, string bordxNo, string payCode, string company, string user, decimal amount, int claims) =>
        session.ExecAsync("cis.BordxInv_SaveBordx", new { bordxNo, payCode, company, user, amount, claims });
}
