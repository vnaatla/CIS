using Cis.Application.Abstractions;
using Cis.Application.Audit;
using Cis.Application.Invoice;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Audit;

/// <summary><see cref="IAuditReviewRepository"/> on cis.Audit_Case and cis.InvAdj_Audit.</summary>
public sealed class AuditReviewRepository : IAuditReviewRepository
{
    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> CaseAsync(IDbSession session, string casNumber) =>
        (await session.RowsAsync("cis.Audit_Case", new { casNumber })).FirstOrDefault();

    /// <inheritdoc />
    public Task UpdateAsync(IDbSession session, string casNumber, AuditReviewRequest r, DateTime auditDate) =>
        session.ExecAsync("cis.InvAdj_Audit", new
        {
            casNumber, query = (r.Query ?? "").Trim(), answer = (r.Answer ?? "").Trim(), discount = (r.CostAvoidance ?? "").Trim(),
            remarks = (r.Remarks ?? "").Trim(), auditBy = (r.AuditBy ?? "").Trim(), auditDate,
        });
}
