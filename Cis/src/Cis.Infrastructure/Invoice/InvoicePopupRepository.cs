using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Infrastructure.Data;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IInvoicePopupRepository"/> on <c>cis.Popup_Icd</c> and <c>cis.Popup_CaseList</c>.</summary>
public sealed class InvoicePopupRepository : IInvoicePopupRepository
{
    /// <inheritdoc />
    public Task<List<IcdRow>> IcdAsync(IDbSession session, string? code, string? desc, string? desc2, bool useOr) =>
        session.QueryAsync<IcdRow>("cis.Popup_Icd", new { code, desc, desc2, useOr });

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> CasesAsync(IDbSession session, string caseType, int by, string text) =>
        session.RowsAsync("cis.Popup_CaseList", new { caseType, by, text });
}
