using Cis.Application.Abstractions;
using Cis.Application.Maintenance.ListReports;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Maintenance;

/// <summary><see cref="IListReportRepository"/> on the <c>cis.ListReport_*</c> procedures.</summary>
public sealed class ListReportRepository : IListReportRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> RowsAsync(IDbSession session, string procedure, string? code) =>
        session.RowsAsync(procedure, new { code });
}
