using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Documents;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Invoice;

/// <summary>The form's search filters (SearchRecord). Dates are dd/mm/yyyy; a date that is not valid is ignored, as in the desktop.</summary>
/// <param name="CasFrom">WS No. from.</param>
/// <param name="CasTo">WS No. to.</param>
/// <param name="Clinic">Clinic combo text ("code - name" or a code).</param>
/// <param name="GroupCompany">Payor (group company).</param>
/// <param name="InvoiceNo">Invoice No.</param>
/// <param name="VisitFrom">Visit date from.</param>
/// <param name="VisitTo">Visit date to.</param>
/// <param name="NoInvoice">Without invoice number (ChkInvNo).</param>
/// <param name="InvoiceFrom">Invoice date from.</param>
/// <param name="InvoiceTo">Invoice date to.</param>
/// <param name="AmountFrom">Invoice amount from.</param>
/// <param name="AmountTo">Invoice amount to.</param>
public record InvUpdateFilters(string? CasFrom, string? CasTo, string? Clinic, string? GroupCompany, string? InvoiceNo, string? VisitFrom, string? VisitTo,
    bool NoInvoice, string? InvoiceFrom, string? InvoiceTo, string? AmountFrom, string? AmountTo);

/// <summary>Repository form of the filters.</summary>
/// <param name="CasFrom">WS No. from.</param>
/// <param name="CasTo">WS No. to.</param>
/// <param name="ClinicCode">Clinic code.</param>
/// <param name="GroupCompany">Group company.</param>
/// <param name="BillNo">Invoice No.</param>
/// <param name="VisitFrom">Visit date from.</param>
/// <param name="VisitTo">Visit date to.</param>
/// <param name="NoInvoice">Without invoice number.</param>
/// <param name="BillFrom">Invoice date from.</param>
/// <param name="BillTo">Invoice date to.</param>
/// <param name="AmountFrom">Amount from.</param>
/// <param name="AmountTo">Amount to.</param>
/// <param name="BillDate">Exact invoice date (the list shown after Update).</param>
public record InvUpdateQuery(string? CasFrom, string? CasTo, string? ClinicCode, string? GroupCompany, string? BillNo, DateTime? VisitFrom, DateTime? VisitTo,
    bool NoInvoice, DateTime? BillFrom, DateTime? BillTo, decimal? AmountFrom, decimal? AmountTo, DateTime? BillDate = null);

/// <summary>Combos.</summary>
/// <param name="Clinics">Panel clinics.</param>
/// <param name="GroupCompanies">Group companies.</param>
public record InvUpdateSetup(List<LookupOption> Clinics, List<string> GroupCompanies);

/// <summary>CmdUpdate_Click.</summary>
/// <param name="CasNumbers">Checked cases.</param>
/// <param name="InvoiceNo">Invoice No. to set.</param>
/// <param name="InvoiceDate">Invoice date (dd/mm/yyyy).</param>
public record InvUpdateRequest(List<string>? CasNumbers, string? InvoiceNo, string? InvoiceDate);

/// <summary>Column lengths (CISCAS).</summary>
public sealed class InvUpdateRequestValidator : RequestValidator<InvUpdateRequest>
{
    /// <summary>Defines the rules.</summary>
    public InvUpdateRequestValidator()
    {
        RuleFor(r => r.InvoiceNo).MaximumLength(50);
        RuleForEach(r => r.CasNumbers).MaximumLength(20);
    }
}

/// <summary>Data access of Invoice Update (<c>cis.InvUpdate_*</c>).</summary>
public interface IInvoiceUpdateRepository
{
    /// <summary>Clinics and group companies.</summary>
    Task<InvUpdateSetup> SetupAsync(IDbSession session);

    /// <summary>VIEW_Invoice_Tracking_2016 rows.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, InvUpdateQuery query);

    /// <summary>UpdateRecord; rows updated.</summary>
    Task<int> SaveAsync(IDbSession session, string casNumber, string invoiceNo, DateTime invoiceDate, DateTime today, string userCode);
}

/// <summary>Invoice > Other > Invoice Update.</summary>
public interface IInvoiceUpdateService
{
    /// <summary>ChkClinic_Click / ChkGrpComp_Click lists.</summary>
    Task<InvUpdateSetup> SetupAsync();

    /// <summary>CmdSearch_Click.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(InvUpdateFilters filters);

    /// <summary>CmdPrint_Click: the list as Excel.</summary>
    Task<FileOutput> ExportAsync(InvUpdateFilters filters);

    /// <summary>CmdUpdate_Click: the message and the cases of the new invoice.</summary>
    Task<(string Message, List<Dictionary<string, object?>> Rows)> SaveAsync(InvUpdateRequest request, string userCode);
}

/// <inheritdoc />
public sealed class InvoiceUpdateService(IDbSessions sessions, IInvoiceUpdateRepository repository, IReportExporter exporter, TimeProvider clock)
    : IInvoiceUpdateService
{
    /// <summary>Desktop texts (spacing kept; "Cheque Amt" is the desktop's wording for a missing invoice number).</summary>
    public const string NotFound = "No record found!", SelectRecord = "Please select the record", InvalidDate = "Invalid Invoice Date!",
        InvoiceFirst = "Please Enter Cheque Amt!", Saved = "Record Saved...", NothingToPrint = "Report cannot be printed without any record.";

    /// <summary>lstInvoiceListing columns (export).</summary>
    public static readonly IReadOnlyList<ReportColumn> Columns =
    [
        new("CASNumber", "Case No"), new("MBMNumber", "Member No"), new("CoverID", "Cover ID"), new("PatientName", "Patient Name"),
        new("DateVisit", "Visit Date", ReportFormat.Date), new("CLNName", "Clinic"), new("BillNo", "Invoice No"), new("BillDate", "Invoice Date", ReportFormat.Date),
        new("INVEnteredDate", "Entered Date", ReportFormat.Date), new("INVEnteredBy", "Entered By"), new("INVAmount", "Invoice Amt", ReportFormat.Money, true),
        new("CLMMCOProcFee", "FFS", ReportFormat.Money), new("CLMFinalApprove", "Final Approve", ReportFormat.Money), new("MBMGroupCompany", "Group Company"),
    ];

    /// <inheritdoc />
    public async Task<InvUpdateSetup> SetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.SetupAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(InvUpdateFilters filters)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, Query(filters));
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<FileOutput> ExportAsync(InvUpdateFilters filters)
    {
        List<Dictionary<string, object?>> rows;
        await using (var s = await sessions.OpenAsync()) rows = await repository.SearchAsync(s, Query(filters));
        if (rows.Count == 0) throw new BusinessRuleException(NothingToPrint);
        return exporter.Export(new ReportDocument("Invoice Update", Columns, rows) { FileName = "InvoiceUpdate" }, ReportFileType.Excel);
    }

    /// <inheritdoc />
    public async Task<(string Message, List<Dictionary<string, object?>> Rows)> SaveAsync(InvUpdateRequest r, string userCode)
    {
        var cases = (r.CasNumbers ?? []).Select(c => c.Trim()).Where(c => c.Length > 0).Distinct().ToList();
        if (cases.Count == 0) throw new BusinessRuleException(SelectRecord);
        var date = Date(r.InvoiceDate) ?? throw new BusinessRuleException(InvalidDate);
        var invoice = r.InvoiceNo ?? "";
        if (invoice.Length == 0) throw new BusinessRuleException(InvoiceFirst);
        var today = clock.GetLocalNow().Date;
        await using var s = await sessions.BeginAsync();
        foreach (var c in cases) await repository.SaveAsync(s, c, invoice, date, today, userCode);
        await s.CommitAsync();
        // The desktop lists the cases of the new invoice number and date afterwards.
        await using var read = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(read, new InvUpdateQuery(null, null, null, null, invoice.Trim(), null, null, false, null, null, null, null, date));
        return (Saved, rows);
    }

    /// <summary>SearchRecord: text filters trimmed (empty = not used); the clinic code is the text before the first "-".</summary>
    public static InvUpdateQuery Query(InvUpdateFilters f)
    {
        static string? T(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        var clinic = T(f.Clinic);
        if (clinic is not null && clinic.IndexOf('-') is var dash and > 0) clinic = clinic[..dash].Trim();
        return new InvUpdateQuery(T(f.CasFrom), T(f.CasTo), clinic, T(f.GroupCompany), T(f.InvoiceNo), Date(f.VisitFrom), Date(f.VisitTo), f.NoInvoice,
            Date(f.InvoiceFrom), Date(f.InvoiceTo), Amount(f.AmountFrom), Amount(f.AmountTo));
    }

    private static DateTime? Date(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static decimal? Amount(string? text) =>
        decimal.TryParse(text?.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : null;
}
