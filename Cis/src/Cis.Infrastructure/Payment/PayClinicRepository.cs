using System.Data;
using Cis.Application.Abstractions;
using Cis.Application.Payment;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Payment;

/// <summary><see cref="IPayClinicRepository"/> on cis.PayClinic_*.</summary>
public sealed class PayClinicRepository : IPayClinicRepository
{
    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PayClinicSearchRequest r, string? pvNo, string? clinic, DateTime?[] dates)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        return session.RowsAsync("cis.PayClinic_Search", new
        {
            pvNo = B(pvNo), company = B(r.Company), clinic, payment = r.Payment, scan = r.Scan, caseFrom = B(r.CaseFrom), caseTo = r.Scan ? null : B(r.CaseTo),
            invoiceNo = B(r.InvoiceNo), blankInvoiceNo = r.BlankInvoiceNo, invFrom = dates[0], invTo = dates[1], blankBatch = r.BlankBatch, opFrom = B(r.OpFrom),
            opTo = B(r.OpTo), blankInvDate = r.BlankInvDate, amtFrom = B(r.AmtFrom), amtTo = B(r.AmtTo), blankAmt = r.BlankAmt, pvFrom = B(r.PvFrom), pvTo = B(r.PvTo),
            blankPv = r.BlankPv, pvDateFrom = dates[2], pvDateTo = dates[3], blankPvDate = r.BlankPvDate, mthFrom = B(r.MthFrom), mthTo = B(r.MthTo),
            blankMth = r.BlankMth, yrFrom = B(r.YrFrom), yrTo = B(r.YrTo), blankYr = r.BlankYr, sunFrom = B(r.SunFrom), sunTo = B(r.SunTo), entity = r.Entity,
        }, 600);
    }

    /// <inheritdoc />
    public async Task<(List<(string Code, string Name, string PvCode)> Banks, List<string> Companies, List<(string Code, string Name)> Clinics)> LookupsAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.PayClinic_Lookups");
        return ((await grid.ReadAsync<(string, string, string)>()).ToList(), (await grid.ReadAsync<string>()).ToList(), (await grid.ReadAsync<(string, string)>()).ToList());
    }

    /// <inheritdoc />
    public async Task<string> NextPvAsync(IDbSession session, string bank, string pvCode)
    {
        var p = new DynamicParameters(new { bank, pvCode });
        p.Add("pvNo", dbType: DbType.String, size: 30, direction: ParameterDirection.Output);
        await session.Sql().Connection.ExecuteAsync("cis.PayClinic_NextPv", p, session.Sql().Transaction, commandType: CommandType.StoredProcedure, commandTimeout: 120);
        return p.Get<string>("pvNo");
    }

    /// <inheritdoc />
    public Task SaveAsync(IDbSession session, PayClinicRow r) =>
        session.ExecAsync("cis.PayClinic_Save", new
        {
            index = r.Index, edit = r.Edit, casNumber = r.CasNumber, pvNo = r.PvNo, pvDate = r.PvDate, chequeDate = r.ChequeDate, chequeNo = r.ChequeNo,
            chequeAmt = r.ChequeAmt, bank = r.Bank, account = r.Account, remark = r.Remark, paidAmt = r.PaidAmt, mth = r.Month, yr = r.Year, status = r.Status, user = r.User,
        });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> PvRowsAsync(IDbSession session, string? pvNo, string? cases) =>
        session.RowsAsync("cis.PayClinic_PvRows", new { pvNo, cases }, 300);
}
