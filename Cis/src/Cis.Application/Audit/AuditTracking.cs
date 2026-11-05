using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Audit;

/// <summary>Audit Tracking filters (FrmAuditInvoiceTracking's search frame). "Blank" flags are the "is empty" check boxes; dates are dd/mm/yyyy
/// (invalid dates are ignored, as the desktop's IsDate checks did); codes are the part before " - ".</summary>
public sealed record AuditTrackingRequest
{
#pragma warning disable CS1591
    public string? CaseFrom { get; init; }
    public string? CaseTo { get; init; }
    public bool CaseBlank { get; init; }
    public string? PayTo { get; init; }
    public bool PayToBlank { get; init; }
    public string? Clinic { get; init; }
    public bool ClinicBlank { get; init; }
    public string? Payor { get; init; }
    public bool PayorBlank { get; init; }
    public string? GroupCompany { get; init; }
    public bool GroupBlank { get; init; }
    public string? PatientName { get; init; }
    public bool NameBlank { get; init; }
    public string? MemberFrom { get; init; }
    public string? MemberTo { get; init; }
    public bool MemberBlank { get; init; }
    public string? InvoiceFrom { get; init; }
    public string? InvoiceTo { get; init; }
    public bool InvoiceBlank { get; init; }
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    public bool BordxBlank { get; init; }
    public string? CaseStatus { get; init; }
    public string? Claimability { get; init; }
    public string? PlanFrom { get; init; }
    public string? PlanTo { get; init; }
    public bool PlanBlank { get; init; }
    public string? BordxDateFrom { get; init; }
    public string? BordxDateTo { get; init; }
    public string? IcFrom { get; init; }
    public string? IcTo { get; init; }
    public bool IcBlank { get; init; }
    public string? VisitFrom { get; init; }
    public string? VisitTo { get; init; }
    public bool VisitBlank { get; init; }
    public string? RegisterFrom { get; init; }
    public string? RegisterTo { get; init; }
    public bool RegisterBlank { get; init; }
    public string? InvoiceDateFrom { get; init; }
    public string? InvoiceDateTo { get; init; }
    public bool InvoiceDateBlank { get; init; }
    public string? EnteredFrom { get; init; }
    public string? EnteredTo { get; init; }
    public bool EnteredBlank { get; init; }
    public string? AmountFrom { get; init; }
    public string? AmountTo { get; init; }
    public bool AmountBlank { get; init; }
    public string? MonthFrom { get; init; }
    public string? MonthTo { get; init; }
    public bool MonthBlank { get; init; }
    public string? YearFrom { get; init; }
    public string? YearTo { get; init; }
    public bool YearBlank { get; init; }
    public string? ClinicType { get; init; }
    public string? OpBatchFrom { get; init; }
    public string? OpBatchTo { get; init; }
    public bool OpBatchBlank { get; init; }
    public string? Ffs { get; init; }
    public bool Omni { get; init; }
    public bool Tims { get; init; }
    public string? Subsidiary { get; init; }
    public string? Department { get; init; }
    public string? OmniRefFrom { get; init; }
    public string? OmniRefTo { get; init; }
    public bool OmniRefBlank { get; init; }
    public string? OmniSent { get; init; }
    public bool Audited { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths and amounts.</summary>
public sealed class AuditTrackingRequestValidator : Common.RequestValidator<AuditTrackingRequest>
{
    /// <summary>Defines the rules.</summary>
    public AuditTrackingRequestValidator()
    {
        RuleFor(r => r.CaseFrom).MaximumLength(20);
        RuleFor(r => r.CaseTo).MaximumLength(20);
        RuleFor(r => r.PatientName).MaximumLength(200);
        RuleFor(r => r.GroupCompany).MaximumLength(200);
        RuleFor(r => r.Subsidiary).MaximumLength(200);
        RuleFor(r => r.Department).MaximumLength(200);
        RuleFor(r => r.AmountFrom).Must(Amount).WithMessage("Invalid amount!");
        RuleFor(r => r.AmountTo).Must(Amount).WithMessage("Invalid amount!");
    }

    private static bool Amount(string? v) => string.IsNullOrWhiteSpace(v) || decimal.TryParse(v, NumberStyles.Number, CultureInfo.InvariantCulture, out _);
}

/// <summary>One listed row: one per medication of a case (cases without medication are not listed, as in Sp_Audittraccking_2018).</summary>
/// <param name="First">First row of its case (the desktop blanked the case columns on the others).</param>
/// <param name="Row">Sp_Audittraccking_2018's columns, plus Diagnosis5 (the typed others diagnosis, else ReasonDesc5).</param>
public record AuditTrackingRow(bool First, Dictionary<string, object?> Row);

/// <summary>Data access of Audit Tracking.</summary>
public interface IAuditTrackingRepository
{
    /// <summary>cis.AuditTrack_Search.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, AuditTrackingRequest filter, AuditTrackingDates dates, decimal? amountFrom, decimal? amountTo);

    /// <summary>cis.AuditTrack_Lookups: subsidiaries and departments of a group company.</summary>
    Task<(List<string> Subsidiaries, List<string> Departments)> LookupsAsync(IDbSession session, string groupCompany);
}

/// <summary>Parsed date filters (null = not used).</summary>
public record AuditTrackingDates(DateTime? BordxFrom, DateTime? BordxTo, DateTime? VisitFrom, DateTime? VisitTo, DateTime? RegisterFrom, DateTime? RegisterTo,
    DateTime? InvoiceFrom, DateTime? InvoiceTo, DateTime? EnteredFrom, DateTime? EnteredTo);

/// <summary>LG &amp; Audit > Audit Tracking (FrmAuditInvoiceTracking).</summary>
public interface IAuditTrackingService
{
    /// <summary>cmdSearch_Click.</summary>
    Task<List<AuditTrackingRow>> SearchAsync(AuditTrackingRequest filter);

    /// <summary>CmdPrintFile_Click: the InvoiceAuditTrackingNEW report.</summary>
    Task<FileOutput> PrintAsync(AuditTrackingRequest filter);

    /// <summary>CboGrpCom_LostFocus.</summary>
    Task<(List<string> Subsidiaries, List<string> Departments)> LookupsAsync(string? groupCompany);
}

/// <inheritdoc />
public sealed class AuditTrackingService(IDbSessions sessions, IAuditTrackingRepository repository, IReportExporter exporter, TimeProvider clock) : IAuditTrackingService
{
    /// <summary>Desktop texts.</summary>
    public const string NoRecord = "No record found!", OmniRefMissing = "Invalid column name 'OMNIRefNo'.";

    /// <summary>The desktop's Excel columns A–Y (InvoiceAuditTrackingNEW.xlt).</summary>
    public static readonly IReadOnlyList<ReportColumn> Columns =
    [
        new("BordxRefNo", "Bordx No"), new("MBMIcBcPp", "IC No"), new("MBMName", "Name"), new("CASNumber", "Case No"),
        new("DateVisit", "Visit Date", ReportFormat.Date), new("ReasonDesc1", "Diagnosis 1"), new("ReasonDesc2", "Diagnosis 2"),
        new("ReasonDesc3", "Diagnosis 3"), new("ReasonDesc4", "Diagnosis 4"), new("Diagnosis5", "Diagnosis 5"), new("MedInjType", "Medication"),
        new("MedInjDays", "Days"), new("MedInjQty", "Qty"), new("MedCharges", "Charges", ReportFormat.Money), new("Procedurefee", "Procedure Fee", ReportFormat.Money),
        new("SPFee", "Consultation Fee", ReportFormat.Money), new("INVAmount", "Invoice Amount", ReportFormat.Money), new("AuditQuery", "Query"),
        new("AuditAnswer", "Answer"), new("AuditRemarks", "Remarks"), new("AuditDiscountReceived", "Discount Received"), new("AuditBy", "Audit By"),
        new("CLNCode", "Clinic Code"), new("CLNName", "Clinic Name"), new("MBMGroupCompany", "Group Company"),
    ];

    private static readonly string[] CaseColumns =
        ["BordxRefNo", "MBMIcBcPp", "MBMName", "CASNumber", "DateVisit", "ReasonDesc1", "ReasonDesc2", "ReasonDesc3", "ReasonDesc4", "Diagnosis5",
         "Procedurefee", "SPFee", "INVAmount", "CasRemark", "AuditQuery", "AuditAnswer", "AuditRemarks", "AuditDiscountReceived", "AuditBy", "AuditDate",
         "CLNCode", "CLNName", "MBMGroupCompany"];

    /// <inheritdoc />
    public async Task<List<AuditTrackingRow>> SearchAsync(AuditTrackingRequest filter)
    {
        var rows = await RowsAsync(filter);
        return rows.Count == 0 ? throw new NotFoundException(NoRecord) : rows;
    }

    /// <inheritdoc />
    public async Task<FileOutput> PrintAsync(AuditTrackingRequest filter)
    {
        var rows = await RowsAsync(filter);
        if (rows.Count == 0) throw new BusinessRuleException("Please search again.!");
        var report = new ReportDocument("INVOICE AUDIT TRACKING", Columns, rows.Select(r => (IReadOnlyDictionary<string, object?>)r.Row).ToList())
        {
            Subtitles = [clock.GetLocalNow().ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)],
            FileName = "InvoiceAuditTracking",
        };
        return exporter.Export(report, ReportFileType.Excel);
    }

    /// <inheritdoc />
    public async Task<(List<string> Subsidiaries, List<string> Departments)> LookupsAsync(string? groupCompany)
    {
        if (string.IsNullOrWhiteSpace(groupCompany)) return ([], []);
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s, groupCompany.Trim());
    }

    private async Task<List<AuditTrackingRow>> RowsAsync(AuditTrackingRequest f)
    {
        // OMNIRefNo is not a column of the report's tables, so the desktop's OMNI Ref filters end in this SQL Server error.
        if (f.OmniRefBlank || !string.IsNullOrWhiteSpace(f.OmniRefFrom) || !string.IsNullOrWhiteSpace(f.OmniRefTo))
            throw new BusinessRuleException(OmniRefMissing);
        var dates = new AuditTrackingDates(D(f.BordxDateFrom), D(f.BordxDateTo), D(f.VisitFrom), D(f.VisitTo), D(f.RegisterFrom), D(f.RegisterTo),
            D(f.InvoiceDateFrom), D(f.InvoiceDateTo), D(f.EnteredFrom), D(f.EnteredTo));
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, f, dates, A(f.AmountFrom), A(f.AmountTo));
        return rows.Select(r =>
        {
            r["Diagnosis5"] = InvoiceAdjustmentService.T(r, "CASDiagnosis5") is { Length: > 0 } d5 ? d5 : r.GetValueOrDefault("ReasonDesc5");
            var first = Convert.ToInt64(r.GetValueOrDefault("RNo") ?? 1L, CultureInfo.InvariantCulture) == 1;
            if (!first) foreach (var c in CaseColumns) r[c] = null;
            return new AuditTrackingRow(first, r);
        }).ToList();
    }

    private static DateTime? D(string? v) => string.IsNullOrWhiteSpace(v) ? null : ClinicInvoiceRules.Date(v);

    private static decimal? A(string? v) => decimal.TryParse(v, NumberStyles.Number, CultureInfo.InvariantCulture, out var a) ? a : null;
}
