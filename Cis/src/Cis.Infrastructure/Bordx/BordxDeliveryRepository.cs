using System.Data;
using Cis.Application.Abstractions;
using Cis.Application.Bordx;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Bordx;

/// <summary><see cref="IBordxDeliveryRepository"/> on cis.BordxDelivery_*.</summary>
public sealed class BordxDeliveryRepository : IBordxDeliveryRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, BordxDeliveryFilter f) =>
        session.RowsAsync("cis.BordxDelivery_Search", new
        {
            kind = f.Kind, company = f.Company, payCode = f.PayCode, scan = f.Scan, bordxFrom = f.BordxFrom, bordxTo = f.BordxTo, dateFrom = f.DateFrom, dateTo = f.DateTo,
            amountFrom = f.AmountFrom, amountTo = f.AmountTo, monthFrom = f.MonthFrom, monthTo = f.MonthTo, yearFrom = f.YearFrom, yearTo = f.YearTo, status = f.Status,
            sentFrom = f.SentFrom, sentTo = f.SentTo, periodMonthFrom = f.PeriodMonthFrom, periodYearFrom = f.PeriodYearFrom, periodMonthTo = f.PeriodMonthTo,
            periodYearTo = f.PeriodYearTo, noPeriod = f.NoPeriod,
        }, 300);

    /// <inheritdoc />
    public async Task<int> UpdateAsync(IDbSession session, string kind, string bordxNo, DateTime? date, string? month = null, string? year = null)
    {
        var p = new DynamicParameters(new { kind, bordxNo, date, month, year });
        p.Add("result", dbType: DbType.Int32, direction: ParameterDirection.Output);
        await session.Sql().Connection.ExecuteAsync("cis.BordxDelivery_Update", p, session.Sql().Transaction, commandType: CommandType.StoredProcedure, commandTimeout: 120);
        return p.Get<int>("result");
    }

    /// <inheritdoc />
    public async Task<(List<string> Companies, List<(string Code, string Name)> Payors)> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.BordxDelivery_Lookups");
        return ((await grid.ReadAsync<string>()).ToList(), (await grid.ReadAsync<(string, string)>()).ToList());
    }
}
