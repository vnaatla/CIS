using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Documents;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Invoice;

/// <summary>The form's filters (SearchRecord). Dates dd/mm/yyyy (invalid ones are ignored); combo values as their first character.</summary>
/// <param name="NoBatch">Without OP batch (ChkBatchNo).</param>
/// <param name="NoBordx">Without bordereaux (ChkBordxNo).</param>
/// <param name="BordxFrom">Bordx ref from.</param>
/// <param name="BordxTo">Bordx ref to.</param>
/// <param name="NoBtbgBordx">Without BTBG bordereaux.</param>
/// <param name="BtbgFrom">BTBG ref from.</param>
/// <param name="BtbgTo">BTBG ref to.</param>
/// <param name="CasFrom">WS No. from.</param>
/// <param name="CasTo">WS No. to.</param>
/// <param name="Clinic">Clinic combo text ("code - name").</param>
/// <param name="GroupCompany">Payor.</param>
/// <param name="InvoiceNo">Invoice No. filter.</param>
/// <param name="VisitFrom">Visit date from.</param>
/// <param name="VisitTo">Visit date to.</param>
/// <param name="NoInvoice">Without invoice number.</param>
/// <param name="InvoiceFrom">Invoice date from.</param>
/// <param name="InvoiceTo">Invoice date to.</param>
/// <param name="AmountFrom">Invoice amount from.</param>
/// <param name="AmountTo">Invoice amount to.</param>
/// <param name="BatchFrom">OP batch from.</param>
/// <param name="BatchTo">OP batch to.</param>
/// <param name="Omni">OMNI status Y.</param>
/// <param name="RegFrom">Register date from.</param>
/// <param name="RegTo">Register date to.</param>
/// <param name="ClinicType">0 Specialised, 1 General Practisional, 2 Others.</param>
/// <param name="Ffs">FFS clinic Y / N.</param>
/// <param name="HasBatch">Invoiced (has OP batch) Y / N.</param>
/// <param name="PayTo">C / M.</param>
/// <param name="Referral">YES / NO.</param>
/// <param name="Entity">0 = not SJ cases, 1 = SJ cases.</param>
public record OpBatchFilters(bool NoBatch, bool NoBordx, string? BordxFrom, string? BordxTo, bool NoBtbgBordx, string? BtbgFrom, string? BtbgTo,
    string? CasFrom, string? CasTo, string? Clinic, string? GroupCompany, string? InvoiceNo, string? VisitFrom, string? VisitTo, bool NoInvoice,
    string? InvoiceFrom, string? InvoiceTo, string? AmountFrom, string? AmountTo, string? BatchFrom, string? BatchTo, bool Omni, string? RegFrom,
    string? RegTo, string? ClinicType, string? Ffs, string? HasBatch, string? PayTo, string? Referral, int Entity = 0);

/// <summary>InvoiceListing result.</summary>
/// <param name="Rows">Cases (Pending "Z" left out).</param>
/// <param name="TotalInvoice">Total invoice amount.</param>
/// <param name="TotalFfs">Total FFS.</param>
/// <param name="MultipleMonths">The visit dates fall in more than one month (tmpMsno &gt; 1).</param>
public record OpBatchList(List<Dictionary<string, object?>> Rows, decimal TotalInvoice, decimal TotalFfs, bool MultipleMonths);

/// <summary>CmdUpdate_Click.</summary>
/// <param name="Filters">The filters of the listed cases.</param>
/// <param name="CasNumbers">Checked cases.</param>
/// <param name="InvoiceNo">Invoice No. (empty = OP batch only).</param>
public record OpBatchRequest(OpBatchFilters? Filters, List<string>? CasNumbers, string? InvoiceNo);

/// <summary>Lengths.</summary>
public sealed class OpBatchRequestValidator : RequestValidator<OpBatchRequest>
{
    /// <summary>Defines the rules.</summary>
    public OpBatchRequestValidator()
    {
        RuleFor(r => r.InvoiceNo).MaximumLength(40);
        RuleForEach(r => r.CasNumbers).MaximumLength(20);
    }
}

/// <summary>Data access of OP Batch/Invoice (<c>cis.OpBatch_*</c>).</summary>
public interface IOpBatchRepository
{
    /// <summary>ClaimBordxModule case 27 rows.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, OpBatchFilters filters);

    /// <summary>The invoice number is used by a case.</summary>
    Task<bool> InvoiceUsedAsync(IDbSession session, string invoiceNo);

    /// <summary>GenerateOPSeq.</summary>
    Task<string> NextBatchAsync(IDbSession session, string opCode);

    /// <summary>UpdateRecord.</summary>
    Task SaveAsync(IDbSession session, string casNumber, string? batchNo, string userCode, DateTime today, string? invoiceNo, DateTime? invoiceDate);
}

/// <summary>Invoice > Other > OP Batch/Invoice.</summary>
public interface IOpBatchService
{
    /// <summary>CmdSearch_Click.</summary>
    Task<OpBatchList> SearchAsync(OpBatchFilters filters);

    /// <summary>CmdPrint_Click.</summary>
    Task<FileOutput> ExportAsync(OpBatchFilters filters);

    /// <summary>CmdUpdate_Click / TxtInvoiceNo_LostFocus.</summary>
    Task<(string Message, string? BatchNo, string? InvoiceNo)> SaveAsync(OpBatchRequest request, string userCode);
}

/// <inheritdoc />
public sealed class OpBatchService(IDbSessions sessions, IOpBatchRepository repository, IReportExporter exporter, TimeProvider clock) : IOpBatchService
{
    /// <summary>Desktop texts (spacing kept).</summary>
    public const string NotFound = "No record found!", SearchFirst = "Please Search Data..!", VisitFirst = "Please select Visit Date.",
        SameClinic = "Above claims must belong to same clinic code.", SameMonth = "All the claims visit date must belong to same Month & Year.",
        InvoiceUsed = "Given Invoice No is already used...!", SelectRecord = "Please select the record", Saved = "Record Saved...",
        NothingToPrint = "Report cannot be printed without any record.", AlreadyDone = "Selected Case Already OP Batch And Invoice No generated..! ",
        IsPending = "Selected Case is Pending..! ", Claimability = "Please check Case Claimability..! ";

    /// <summary>lstInvoiceListing columns (export).</summary>
    public static readonly IReadOnlyList<ReportColumn> Columns =
    [
        new("CASNumber", "Case No"), new("MBMNumber", "Member No"), new("CoverID", "Cover ID"), new("PatientName", "Patient Name"),
        new("DateVisit", "Visit Date", ReportFormat.Date), new("MBMGroupCompany", "Group Company"), new("INVAmount", "Invoice Amt", ReportFormat.Money, true),
        new("CLMMCOProcFee", "FFS", ReportFormat.Money, true), new("CLMFinalApprove", "Final Approve", ReportFormat.Money), new("BillNo", "Invoice No"),
        new("MBMDepartment", "Department"), new("CLNName", "Clinic"), new("BillDate", "Invoice Date", ReportFormat.Date), new("INVEnteredDate", "Entered Date"),
        new("INVEnteredBy", "Entered By"), new("OPBatchNo", "OP Batch No"), new("PNO", "PV No"), new("PDATE", "PV Date", ReportFormat.Date),
        new("PCNO", "Cheque No"), new("ReasonDesc1", "Reason 1"), new("ReasonDesc2", "Reason 2"), new("ReasonDesc3", "Reason 3"),
        new("ReasonDesc4", "Reason 4"), new("ReasonDesc5", "Reason 5"), new("CASDiagnosis5", "Diagnosis"), new("MCDays", "MC Days"),
        new("CLMCoPayCLNColl", "Co-Pay Collected", ReportFormat.Money), new("CLMFinalApprove1", "Balance", ReportFormat.Money), new("ICNo", "IC No"),
        new("OPNO", "OP"), new("PType", "Pay Type"), new("OPBatchDate", "OP Batch Date", ReportFormat.Date), new("LastUpdateDate", "Last Update"),
        new("Pending", "Pending"), new("Claimability", "Claimability"),
    ];

    /// <inheritdoc />
    public async Task<OpBatchList> SearchAsync(OpBatchFilters filters)
    {
        var list = await List(filters);
        return list.Rows.Count == 0 ? throw new NotFoundException(NotFound) : list;
    }

    /// <inheritdoc />
    public async Task<FileOutput> ExportAsync(OpBatchFilters filters)
    {
        var list = await List(filters);
        if (list.Rows.Count == 0) throw new BusinessRuleException(NothingToPrint);
        return exporter.Export(new ReportDocument("OP Batch / Invoice", Columns, list.Rows) { FileName = "OPBatchInvoice" }, ReportFileType.Excel);
    }

    /// <inheritdoc />
    public async Task<(string Message, string? BatchNo, string? InvoiceNo)> SaveAsync(OpBatchRequest r, string userCode)
    {
        var filters = r.Filters ?? throw new BusinessRuleException(SearchFirst);
        var list = await List(filters);
        var invoice = (r.InvoiceNo ?? "").Trim();
        string? number = null;
        DateTime? invoiceDate = null;
        if (invoice.Length > 0)
        {
            if (list.Rows.Count == 0) throw new BusinessRuleException(SearchFirst);
            if (Date(filters.VisitFrom) is null || Date(filters.VisitTo) is null) throw new BusinessRuleException(VisitFirst);
            if (string.IsNullOrEmpty(filters.Clinic)) throw new BusinessRuleException(SameClinic);
            if (list.MultipleMonths) throw new BusinessRuleException(SameMonth);
            var last = (DateTime)list.Rows[^1]["DateVisit"]!;
            invoiceDate = new DateTime(last.Year, last.Month, 1).AddMonths(1).AddDays(-1);
            number = InvoiceNumber(invoice, filters.Clinic);
        }
        var listed = list.Rows.ToDictionary(x => T(x, "CASNumber"));
        var chosen = (r.CasNumbers ?? []).Select(c => c.Trim()).Where(listed.ContainsKey).Distinct().ToHashSet();
        // In list order, as the desktop walks its list.
        var rows = list.Rows.Where(x => chosen.Contains(T(x, "CASNumber"))).ToList();
        foreach (var row in rows)
            if (CheckMessage(row) is { } message) throw new BusinessRuleException(message);

        await using var s = await sessions.BeginAsync();
        if (number is not null && await repository.InvoiceUsedAsync(s, number)) throw new BusinessRuleException(InvoiceUsed);
        if (rows.Count == 0) throw new BusinessRuleException(SelectRecord);
        string? batch = null;
        var today = clock.GetLocalNow().Date;
        foreach (var row in rows)
        {
            var noBatch = T(row, "OPBatchNo").Length == 0;
            if (noBatch) batch ??= await repository.NextBatchAsync(s, "OC");
            if (noBatch || number is not null)
                await repository.SaveAsync(s, T(row, "CASNumber"), noBatch ? batch : null, userCode, today, number, invoiceDate);
        }
        await s.CommitAsync();
        return (Saved, batch, number);
    }

    /// <summary>lstInvoiceListing_ItemCheck: why a case cannot be checked (null = it can).</summary>
    public static string? CheckMessage(Dictionary<string, object?> row) =>
        T(row, "OPBatchNo").Length > 0 && T(row, "BillNo").Length > 0 ? AlreadyDone
        : T(row, "Pending") == "Y" ? IsPending
        : T(row, "Claimability") != "A" ? Claimability
        : null;

    /// <summary>tmpno: the invoice number followed by "-" and the first four characters of the clinic combo.</summary>
    public static string InvoiceNumber(string invoice, string clinic) => invoice + "-" + (clinic.Length > 4 ? clinic[..4] : clinic);

    /// <summary>InvoiceListing's month check: a visit in another month than the first row's.</summary>
    public static bool MultipleMonths(IEnumerable<DateTime?> visits)
    {
        string? first = null;
        foreach (var v in visits)
        {
            var month = v?.ToString("MM/yyyy", CultureInfo.InvariantCulture) ?? "";
            if (first is null) first = month;
            else if (month != first) return true;
        }
        return false;
    }

    private async Task<OpBatchList> List(OpBatchFilters filters)
    {
        await using var s = await sessions.OpenAsync();
        var rows = (await repository.SearchAsync(s, filters)).Where(x => T(x, "Pending") != "Z").ToList();
        return new OpBatchList(rows, rows.Sum(x => N(x, "INVAmount")), rows.Sum(x => N(x, "CLMMCOProcFee")),
            MultipleMonths(rows.Select(x => x.GetValueOrDefault("DateVisit") as DateTime?)));
    }

    /// <summary>dd/mm/yyyy, or null.</summary>
    public static DateTime? Date(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static string T(Dictionary<string, object?> row, string key) =>
        row.TryGetValue(key, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(Dictionary<string, object?> row, string key) => row.GetValueOrDefault(key) is decimal d ? d : 0;
}
