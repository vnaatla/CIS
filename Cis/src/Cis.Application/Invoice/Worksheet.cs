using Cis.Application.Abstractions;
using Cis.Domain.Errors;

namespace Cis.Application.Invoice;

/// <summary>Data access of the worksheet (<c>cis.Worksheet_Lookups</c>).</summary>
public interface IWorksheetRepository
{
    /// <summary>The values frmWorksheet looks up for the invoice row.</summary>
    Task<WorksheetLookups> LookupsAsync(IDbSession session, WorksheetInput input);
}

/// <summary>The worksheet pop-up (frmWorksheet), opened from the invoice and adjustment screens.</summary>
public interface IWorksheetService
{
    /// <summary>Form_Load: the worksheet of the invoice row.</summary>
    Task<WorksheetResult> CalculateAsync(WorksheetInput input, string userCode);
}

/// <inheritdoc />
public sealed class WorksheetService(IDbSessions sessions, IWorksheetRepository repository, TimeProvider clock) : IWorksheetService
{
    /// <summary>The grid has 31 benefit rows.</summary>
    public const int Rows = 31;

    /// <inheritdoc />
    public async Task<WorksheetResult> CalculateAsync(WorksheetInput input, string userCode)
    {
        if (input.Items.Count > Rows) throw new BusinessRuleException($"The worksheet has {Rows} benefit rows.");
        WorksheetLookups db;
        await using (var s = await sessions.OpenAsync()) db = await repository.LookupsAsync(s, input);
        return WorksheetEngine.Calculate(input, db, clock.GetLocalNow().Date, userCode);
    }
}
