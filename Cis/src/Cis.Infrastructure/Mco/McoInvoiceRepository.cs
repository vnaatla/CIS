using Cis.Application.Abstractions;
using Cis.Application.Mco;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Mco;

/// <summary><see cref="IMcoInvoiceRepository"/> on the cis.McoInv_* procedures.</summary>
public sealed class McoInvoiceRepository : IMcoInvoiceRepository
{
    private static object P(McoInvoiceFilter f) => new
    {
        mode = f.Mode, bordxDate = f.BordxDate, batchNo = f.BatchNo, from = f.From, to = f.To, invoiceNo = f.InvoiceNo, version = f.Version, payor = f.Payor,
        groupCode = f.GroupCode, department = f.Department, reprint = f.Reprint,
    };

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> PrincipalsAsync(IDbSession session, McoInvoiceFilter f) => session.RowsAsync("cis.McoInv_Principals", P(f), 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SuppsAsync(IDbSession session, McoInvoiceFilter f, string? member) =>
        session.RowsAsync("cis.McoInv_Supps", new
        {
            mode = f.Mode, bordxDate = f.BordxDate, batchNo = f.BatchNo, from = f.From, to = f.To, invoiceNo = f.InvoiceNo, version = f.Version, payor = f.Payor,
            groupCode = f.GroupCode, department = f.Department, reprint = f.Reprint, member,
        }, 300);

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SummaryAsync(IDbSession session, McoInvoiceFilter f) => session.RowsAsync("cis.McoInv_Summary", P(f), 300);

    /// <inheritdoc />
    public async Task<Dictionary<string, object?>?> CompanyAsync(IDbSession session, string code, bool byPayor) =>
        (await session.RowsAsync("cis.McoInv_Company", new { code, byPayor })).FirstOrDefault();

    /// <inheritdoc />
    public Task StampAsync(IDbSession session, string member, string cover, string invoiceNo, string user) =>
        session.ExecAsync("cis.McoInv_Stamp", new { member, cover, invoiceNo, user });

    /// <inheritdoc />
    public async Task<List<string>> DepartmentsAsync(IDbSession session, string company) =>
        (await session.RowsAsync("cis.McoInv_Departments", new { company })).Select(r => Convert.ToString(r["Department"])!).ToList();
}
