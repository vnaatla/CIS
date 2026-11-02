using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Domain.Errors;

namespace Cis.Application.Invoice;

/// <summary>A treatment reason (ICD) of frmSearchICD.</summary>
/// <param name="ReasonCode">Code.</param>
/// <param name="ReasonDesc">Description.</param>
public record IcdRow(string ReasonCode, string? ReasonDesc);

/// <summary>A case of frmCasList.</summary>
/// <param name="CasNumber">Case No.</param>
/// <param name="MbmNumber">Member No.</param>
/// <param name="CoverId">Cover ID.</param>
/// <param name="Name">Patient (blank → "Member may has been Deleted!").</param>
/// <param name="PolicyNo">Policy No.</param>
/// <param name="VisitDate">Visit date.</param>
/// <param name="Status">Case status.</param>
/// <param name="Clinic">"code - name" (panel clinics) or the code.</param>
/// <param name="Ic">IC.</param>
public record CaseListRow(string CasNumber, string MbmNumber, string CoverId, string Name, string PolicyNo, DateTime? VisitDate, string Status, string Clinic,
    string Ic);

/// <summary>Data access of the invoice pop-ups.</summary>
public interface IInvoicePopupRepository
{
    /// <summary>TreatmentReason search.</summary>
    Task<List<IcdRow>> IcdAsync(IDbSession session, string? code, string? desc, string? desc2, bool useOr);

    /// <summary>VIEW_Cas_MBM search.</summary>
    Task<List<Dictionary<string, object?>>> CasesAsync(IDbSession session, string caseType, int by, string text);
}

/// <summary>Invoice pop-ups: ICD search and case list.</summary>
public interface IInvoicePopupService
{
    /// <summary>frmSearchICD cmdSearch_Click.</summary>
    Task<List<IcdRow>> IcdAsync(string? code, string? desc, string? desc2, string? criteria);

    /// <summary>frmCasList cmdSearch_Click.</summary>
    Task<List<CaseListRow>> CasesAsync(bool dental, int by, string? text);
}

/// <inheritdoc />
public sealed class InvoicePopupService(IDbSessions sessions, IInvoicePopupRepository repository) : IInvoicePopupService
{
    /// <summary>Desktop texts.</summary>
    public const string IcdNotFound = "Record not found", TooShort = "Search String must be at least 3 characters!", CaseNotFound = "No Record Available",
        Deleted = "Member may has been Deleted!";

    /// <inheritdoc />
    public async Task<List<IcdRow>> IcdAsync(string? code, string? desc, string? desc2, string? criteria)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.IcdAsync(s, T(code), T(desc), T(desc2), string.Equals(criteria?.Trim(), "OR", StringComparison.OrdinalIgnoreCase));
        return rows.Count == 0 ? throw new NotFoundException(IcdNotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<List<CaseListRow>> CasesAsync(bool dental, int by, string? text)
    {
        // The desktop checks the length before trimming.
        if ((text ?? "").Length < 3) throw new BusinessRuleException(TooShort);
        await using var s = await sessions.OpenAsync();
        var rows = (await repository.CasesAsync(s, dental ? "D" : "C", by is >= 1 and <= 5 ? by : 3, text!.Trim())).Select(Row).ToList();
        return rows.Count == 0 ? throw new NotFoundException(CaseNotFound) : rows;
    }

    /// <summary>ListCases' columns.</summary>
    public static CaseListRow Row(Dictionary<string, object?> r)
    {
        var cov = S(r, "CoverID");
        var name = cov != "00" ? S(r, "MBMCName") : S(r, "MBMName");
        var code = S(r, "ClinicCode");
        var clinic = code.Length > 0 && S(r, "CLNName") is { Length: > 0 } cln ? $"{code} - {cln}" : code;
        return new CaseListRow(S(r, "CASNumber"), S(r, "MBMNumber"), cov, name.Length > 0 ? name : Deleted, S(r, "MBMPolicyNo"),
            r.GetValueOrDefault("DateVisit") as DateTime?, S(r, "CASStatus"), clinic, S(r, "MBMIcBcPp"));
    }

    private static string? T(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string S(Dictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";
}
