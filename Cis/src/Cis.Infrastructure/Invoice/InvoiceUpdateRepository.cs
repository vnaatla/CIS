using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Application.Lookups;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IInvoiceUpdateRepository"/> on the <c>cis.InvUpdate_*</c> procedures.</summary>
public sealed class InvoiceUpdateRepository : IInvoiceUpdateRepository
{
    /// <inheritdoc />
    public async Task<InvUpdateSetup> SetupAsync(IDbSession session)
    {
        using var grid = await session.MultipleAsync("cis.InvUpdate_Lookups");
        var clinics = (await grid.ReadAsync<LookupOption>()).ToList();
        var groups = (await grid.ReadAsync<string>()).ToList();
        return new InvUpdateSetup(clinics, groups);
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, InvUpdateQuery q) =>
        session.RowsAsync("cis.InvUpdate_Search", new
        {
            casFrom = q.CasFrom, casTo = q.CasTo, clinicCode = q.ClinicCode, groupCompany = q.GroupCompany, billNo = q.BillNo, visitFrom = q.VisitFrom,
            visitTo = q.VisitTo, noInvoice = q.NoInvoice, billFrom = q.BillFrom, billTo = q.BillTo, amountFrom = q.AmountFrom, amountTo = q.AmountTo,
            billDate = q.BillDate,
        });

    /// <inheritdoc />
    public Task<int> SaveAsync(IDbSession session, string casNumber, string invoiceNo, DateTime invoiceDate, DateTime today, string userCode) =>
        session.ScalarAsync<int>("cis.InvUpdate_Save", new { casNumber, invoiceNo, invoiceDate, today, userCode });
}
