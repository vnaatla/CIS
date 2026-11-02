using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Documents;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Claim;

/// <summary>The form's filters (SearchRecord). Dates dd/mm/yyyy (invalid ones are ignored); combos as their first character.</summary>
/// <param name="Scan">Scan mode (ChkScan): one case number.</param>
/// <param name="CaseNo">Case No. (from).</param>
/// <param name="CaseNo2">Case No. to (not in scan mode).</param>
/// <param name="CaseStatus">0 Open, 1 Delete, 2 Open &amp; Close.</param>
/// <param name="Claimability">A / R.</param>
/// <param name="PassBy">Passed by.</param>
/// <param name="Status">P Passed / N Not.</param>
/// <param name="NoVisit">Date of admission empty.</param>
/// <param name="NoRegister">Date of register empty.</param>
/// <param name="NoPass">Date passed empty.</param>
/// <param name="RegFrom">Register date from.</param>
/// <param name="RegTo">Register date to.</param>
/// <param name="VisitFrom">Admission (visit) date from.</param>
/// <param name="VisitTo">Admission (visit) date to.</param>
/// <param name="PassFrom">Pass date from.</param>
/// <param name="PassTo">Pass date to.</param>
public record PassFilters(bool Scan, string? CaseNo, string? CaseNo2, string? CaseStatus, string? Claimability, string? PassBy, string? Status, bool NoVisit,
    bool NoRegister, bool NoPass, string? RegFrom, string? RegTo, string? VisitFrom, string? VisitTo, string? PassFrom, string? PassTo);

/// <summary>A row of lstCaseTrackList.</summary>
/// <param name="Passed">"Passed" / "Not".</param>
/// <param name="CasNumber">Case No.</param>
/// <param name="MbmNumber">Member No.</param>
/// <param name="PatientName">Patient.</param>
/// <param name="RegisterDate">Register date.</param>
/// <param name="VisitDate">Admission date.</param>
/// <param name="PassDate">Pass date.</param>
/// <param name="NoOfDay">No. of days ("0" shows "1").</param>
/// <param name="OpBatchNo">OP batch.</param>
/// <param name="PassBy">Passed by.</param>
/// <param name="ClinicName">Clinic.</param>
/// <param name="ClaimStatus">Claim status.</param>
public record PassRow(string Passed, string CasNumber, string MbmNumber, string PatientName, DateTime? RegisterDate, DateTime? VisitDate, DateTime? PassDate,
    string NoOfDay, string OpBatchNo, string PassBy, string ClinicName, string ClaimStatus);

/// <summary>CmdPass_Click.</summary>
/// <param name="Filters">The filters of the listed cases.</param>
/// <param name="CasNumbers">Checked cases.</param>
public record PassRequest(PassFilters? Filters, List<string>? CasNumbers);

/// <summary>Lengths (ClaimFileTransfer.CASNumber nvarchar(10)).</summary>
public sealed class PassRequestValidator : RequestValidator<PassRequest>
{
    /// <summary>Defines the rules.</summary>
    public PassRequestValidator() => RuleForEach(r => r.CasNumbers).MaximumLength(10);
}

/// <summary>Data access of Pass to Claim (<c>cis.Pass2Claim_*</c>).</summary>
public interface IPassToClaimRepository
{
    /// <summary>SearchCASToken.</summary>
    Task<List<string>> UsersAsync(IDbSession session);

    /// <summary>VIEW_ClaimPass2Claim rows; <paramref name="casNumbers"/> limits them to the cases just passed.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, PassFilters filters, string? casNumbers);

    /// <summary>ClaimPass2ClaimUpdate.</summary>
    Task PassAsync(IDbSession session, string casNumber, string userCode, DateTime today);
}

/// <summary>Claim > Pass to Claim.</summary>
public interface IPassToClaimService
{
    /// <summary>chkPass_Click.</summary>
    Task<List<string>> UsersAsync();

    /// <summary>cmdSearch_Click.</summary>
    Task<List<PassRow>> SearchAsync(PassFilters filters);

    /// <summary>CmdPrintFile_Click.</summary>
    Task<FileOutput> ExportAsync(PassFilters filters);

    /// <summary>CmdPass_Click: the message and the passed cases.</summary>
    Task<(string Message, List<PassRow> Rows)> PassAsync(PassRequest request, string userCode);
}

/// <inheritdoc />
public sealed class PassToClaimService(IDbSessions sessions, IPassToClaimRepository repository, IReportExporter exporter, TimeProvider clock)
    : IPassToClaimService
{
    /// <summary>Desktop texts (spacing kept).</summary>
    public const string CaseFirst = "Please Enter Case No", NotFound = "No record found!", ListFirst = "Please select a payor before proceeding!",
        Updated = "Record updated...", NothingToPrint = "Report cannot be printed without any record.", NoRegister = "Date of Register is Empty.",
        NoAdmission = "Date of Admitted is Empty.", NoClinic = "Clinic Name is Empty.", SentOut = "File has been send out.";

    /// <summary>The list's columns (export).</summary>
    public static readonly IReadOnlyList<ReportColumn> Columns =
    [
        new("Passed", "Status"), new("CasNumber", "Case No"), new("MbmNumber", "Member No"), new("PatientName", "Patient Name"),
        new("RegisterDate", "Date of Register", ReportFormat.Date), new("VisitDate", "Date of Admission", ReportFormat.Date),
        new("PassDate", "Date Passed", ReportFormat.Date), new("NoOfDay", "No of Day"), new("OpBatchNo", "OP Batch No"), new("PassBy", "Pass By"),
        new("ClinicName", "Clinic"), new("ClaimStatus", "Claim Status"),
    ];

    /// <inheritdoc />
    public async Task<List<string>> UsersAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.UsersAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<PassRow>> SearchAsync(PassFilters filters)
    {
        if (filters.Scan && string.IsNullOrWhiteSpace(filters.CaseNo)) throw new BusinessRuleException(CaseFirst);
        var rows = await List(filters, null);
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<FileOutput> ExportAsync(PassFilters filters)
    {
        var rows = await List(filters, null);
        if (rows.Count == 0) throw new BusinessRuleException(NothingToPrint);
        var dicts = rows.Select(r => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["Passed"] = r.Passed, ["CasNumber"] = r.CasNumber, ["MbmNumber"] = r.MbmNumber, ["PatientName"] = r.PatientName, ["RegisterDate"] = r.RegisterDate,
            ["VisitDate"] = r.VisitDate, ["PassDate"] = r.PassDate, ["NoOfDay"] = r.NoOfDay, ["OpBatchNo"] = r.OpBatchNo, ["PassBy"] = r.PassBy,
            ["ClinicName"] = r.ClinicName, ["ClaimStatus"] = r.ClaimStatus,
        }).ToList();
        return exporter.Export(new ReportDocument("Pass to Claim", Columns, dicts) { FileName = "PassToClaim" }, ReportFileType.Excel);
    }

    /// <inheritdoc />
    public async Task<(string Message, List<PassRow> Rows)> PassAsync(PassRequest r, string userCode)
    {
        var filters = r.Filters ?? throw new BusinessRuleException(ListFirst);
        var listed = (await List(filters, null)).ToDictionary(x => x.CasNumber, StringComparer.OrdinalIgnoreCase);
        if (listed.Count == 0) throw new BusinessRuleException(ListFirst);
        var chosen = (r.CasNumbers ?? []).Select(c => c.Trim()).Where(listed.ContainsKey).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        foreach (var c in chosen)
            if (CheckMessage(listed[c]) is { } message) throw new BusinessRuleException(message);
        var today = clock.GetLocalNow().Date;
        // The desktop commits even when nothing was checked (and relists nothing new).
        await using (var s = await sessions.BeginAsync())
        {
            foreach (var c in chosen) await repository.PassAsync(s, listed[c].CasNumber, userCode, today);
            await s.CommitAsync();
        }
        return (Updated, chosen.Count == 0 ? [] : await List(filters, string.Join(",", chosen.Select(c => listed[c].CasNumber))));
    }

    /// <summary>lstCaseTrackList_ItemCheck: why a case cannot be checked (null = it can).</summary>
    public static string? CheckMessage(PassRow r) =>
        r.RegisterDate is null ? NoRegister : r.VisitDate is null ? NoAdmission : r.ClinicName.Length == 0 ? NoClinic : r.ClaimStatus == "1" ? SentOut : null;

    /// <summary>DischargeListing: rows whose cover id equals the member's cover id.</summary>
    public static PassRow? Row(Dictionary<string, object?> r)
    {
        if (!string.Equals(S(r, "CoverID", false), S(r, "MBMCoverID", false), StringComparison.Ordinal)) return null;
        var status = S(r, "ClaimStatus");
        var days = S(r, "NoofDay");
        return new PassRow(status == "P" ? "Passed" : "Not", S(r, "CASNumber"), S(r, "MBMNumber"), S(r, "PatientName"), r.GetValueOrDefault("RegisterDate") as DateTime?,
            r.GetValueOrDefault("DateVisit") as DateTime?, r.GetValueOrDefault("ClaimPassDate") as DateTime?, days == "0" ? "1" : days, S(r, "OPBatchNo"),
            S(r, "ClaimPassBy"), S(r, "CLNName"), S(r, "ClaimStatus", false));
    }

    private async Task<List<PassRow>> List(PassFilters filters, string? casNumbers)
    {
        await using var s = await sessions.OpenAsync();
        return (await repository.SearchAsync(s, filters, casNumbers)).Select(Row).OfType<PassRow>().ToList();
    }

    /// <summary>dd/mm/yyyy, or null.</summary>
    public static DateTime? Date(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static string S(Dictionary<string, object?> r, string k, bool trim = true) =>
        r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture) is var t && trim ? t!.Trim() : t! : "";
}
