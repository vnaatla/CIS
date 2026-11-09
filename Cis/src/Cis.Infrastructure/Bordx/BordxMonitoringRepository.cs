using Cis.Application.Abstractions;
using Cis.Application.Bordx;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Bordx;

/// <summary><see cref="IBordxMonitoringRepository"/> on cis.BordxControl_Search, cis.BordxOs_Search, cis.BordxWks_Search.</summary>
public sealed class BordxMonitoringRepository : IBordxMonitoringRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> ControlAsync(IDbSession session, BordxControlRequest r, DateTime? from, DateTime? to, string? payTo) =>
        session.RowsAsync("cis.BordxControl_Search", new
        {
            company = r.Company, bordxFrom = r.BordxFrom, bordxTo = r.BordxTo, dateFrom = from, dateTo = to, sent = r.Sent, returned = r.Returned, paid = r.Paid,
            subsidiary = r.Subsidiary, payTo, entity = r.Entity,
        }, 600);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> OutstandingAsync(IDbSession session, string? company) =>
        session.RowsAsync("cis.BordxOs_Search", new { company }, 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> WorksheetsAsync(IDbSession session, BordxWksRequest r, string? payCode, DateTime? recFrom, DateTime? recTo) =>
        session.RowsAsync("cis.BordxWks_Search", new
        {
            company = r.Company, payCode, bordxFrom = r.BordxFrom, bordxTo = r.BordxTo, caseFrom = r.CaseFrom, caseTo = r.CaseTo, receiptFrom = r.ReceiptFrom,
            receiptTo = r.ReceiptTo, recFrom, recTo, noRecDate = r.NoRecDate,
        }, 300);

    /// <inheritdoc />
    public async Task<(List<string> Companies, List<(string Code, string Name)> Payors, List<string> Subsidiaries)> LookupsAsync(IDbSession session, string? company)
    {
        using var grid = await session.MultipleAsync("cis.BordxMonitoring_Lookups", new { company });
        return ((await grid.ReadAsync<string>()).ToList(), (await grid.ReadAsync<(string, string)>()).ToList(), (await grid.ReadAsync<string>()).ToList());
    }
}
