using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Infrastructure.Data;
using Dapper;

namespace Cis.Infrastructure.Invoice;

/// <summary><see cref="IInvoiceEnquiryRepository"/> on the <c>cis.InvEnq_*</c> procedures.</summary>
public sealed class InvoiceEnquiryRepository : IInvoiceEnquiryRepository
{
    private static async Task<List<Dictionary<string, object?>>> Read(SqlMapper.GridReader g) =>
        (await g.ReadAsync()).Select(r => ((IDictionary<string, object?>)r).ToDictionary(x => x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase)).ToList();

    /// <inheritdoc />
    public async Task<(Dictionary<string, object?>? Case, Dictionary<string, object?>? Fees)> CaseAsync(IDbSession session, string casNumber)
    {
        using var g = await session.MultipleAsync("cis.InvEnq_Case", new { casNumber });
        var c = (await Read(g)).FirstOrDefault();
        var f = (await Read(g)).FirstOrDefault();
        return (c, f);
    }

    /// <inheritdoc />
    public Task<string?> StepAsync(IDbSession session, string casNumber, bool next) =>
        session.ScalarAsync<string>("cis.InvEnq_Step", new { casNumber, next });

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>?>> WorksheetAsync(IDbSession session, string casNumber)
    {
        using var g = await session.MultipleAsync("cis.InvEnq_Worksheet", new { casNumber });
        var list = new List<Dictionary<string, object?>?>();
        for (var i = 0; i < 8; i++) list.Add((await Read(g)).FirstOrDefault());
        return list;
    }

    /// <inheritdoc />
    public Task<List<Dictionary<string, object?>>> BordxAsync(IDbSession session, string casNumber) => session.RowsAsync("cis.InvEnq_Bordx", new { casNumber });

    /// <inheritdoc />
    public async Task<(List<Dictionary<string, object?>> Payments, List<Dictionary<string, object?>> Receipts)> PaymentsAsync(IDbSession session, string casNumber)
    {
        using var g = await session.MultipleAsync("cis.InvEnq_Payments", new { casNumber });
        return (await Read(g), await Read(g));
    }

    /// <inheritdoc />
    public async Task<(List<Dictionary<string, object?>> Requests, List<Dictionary<string, object?>> Medicines)> MedRequestsAsync(IDbSession session, string casNumber)
    {
        using var g = await session.MultipleAsync("cis.InvEnq_MedRequests", new { casNumber });
        var requests = await Read(g);
        var medicines = g.IsConsumed ? [] : await Read(g);
        return (requests, medicines);
    }
}
