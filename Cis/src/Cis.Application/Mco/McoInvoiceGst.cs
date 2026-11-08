using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Mco;

/// <summary>MCO Invoice GST search (frmMCOFeeGST): the MCO Invoice filters plus the subsidiary.</summary>
public sealed record McoGstRequest
{
#pragma warning disable CS1591
    public int? Criteria { get; init; }
    public string? BordxDate { get; init; }
    public string? BatchNo { get; init; }
    public string? From { get; init; }
    public string? To { get; init; }
    public string? InvoiceNo { get; init; }
    /// <summary>0 Invoice - by Payor, 1 Invoice - by Group Company.</summary>
    public int? Version { get; init; }
    public string? Payor { get; init; }
    public string? GroupCompany { get; init; }
    public string? Subsidiary { get; init; }
    /// <summary>0 Print New, 1 Reprint.</summary>
    public int? PrintOption { get; init; }
    public bool Print { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class McoGstRequestValidator : Common.RequestValidator<McoGstRequest>
{
    /// <summary>Defines the rules.</summary>
    public McoGstRequestValidator()
    {
        RuleFor(r => r.InvoiceNo).MaximumLength(30);
        RuleFor(r => r.Payor).MaximumLength(200);
        RuleFor(r => r.GroupCompany).MaximumLength(300);
        RuleFor(r => r.Subsidiary).MaximumLength(200);
    }
}

/// <summary>The member list (principals and covered persons in one view).</summary>
public record McoGstResult(List<Dictionary<string, object?>> Rows, int Principals, int Supplementaries, string InvoiceNo, bool MultipleInvoices);

/// <summary>cis.McoGst_* parameters.</summary>
public record McoGstFilter(int Mode, DateTime? BordxDate, int? BatchNo, DateTime? From, DateTime? To, string? InvoiceNo, string? GroupCode, string? Subsidiary,
    string? Payor);

/// <summary>Data access.</summary>
public interface IMcoGstRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> ListAsync(IDbSession session, McoGstFilter filter);
    Task<List<Dictionary<string, object?>>> SummaryAsync(IDbSession session, McoGstFilter filter);
    Task<List<Dictionary<string, object?>>> DetailsAsync(IDbSession session, McoGstFilter filter);
    Task<List<string>> SubsidiariesAsync(IDbSession session, string company);
#pragma warning restore CS1591
}

/// <summary>MCO Fees > MCO Invoice GST (frmMCOFeeGST).</summary>
public interface IMcoGstService
{
    /// <summary>BindSubsidary.</summary>
    Task<List<string>> SubsidiariesAsync(string? groupCompany);

    /// <summary>cmdINVSearch_Click.</summary>
    Task<McoGstResult> SearchAsync(McoGstRequest request);

    /// <summary>cmdINVPrint_Click (number, stamp, PrintInvoice) / cmdINVPrintPreview_Click (PrintPreview): summary and detail sheets.</summary>
    Task<FileOutput> PrintAsync(McoGstRequest request, string user);
}

/// <inheritdoc />
public sealed class McoGstService(IDbSessions sessions, IMcoGstRepository repository, IMcoInvoiceRepository invoices, IMcoListRepository numbers,
    IReportExporter exporter) : IMcoGstService
{
    /// <inheritdoc />
    public async Task<List<string>> SubsidiariesAsync(string? groupCompany)
    {
        var parts = (groupCompany ?? "").Split('-');
        if (parts.Length < 2 || string.IsNullOrWhiteSpace(parts[1])) return [];
        await using var s = await sessions.OpenAsync();
        return await repository.SubsidiariesAsync(s, parts[1].Trim());
    }

    /// <inheritdoc />
    public async Task<McoGstResult> SearchAsync(McoGstRequest r)
    {
        Check(r);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.ListAsync(s, ListFilter(r));
        if (rows.Count == 0) throw new NotFoundException(McoInvoiceService.NoRecord);
        var invoices = rows.Select(x => T(x, "MBMCInvoiceNo")).Where(x => x.Length > 0).ToList();
        var princ = rows.Count(x => T(x, "RelCode") == "P");
        return new McoGstResult(rows, princ, rows.Count - princ, invoices.Count > 0 ? invoices[^1] : "", invoices.Distinct().Count() > 1);
    }

    /// <inheritdoc />
    public async Task<FileOutput> PrintAsync(McoGstRequest r, string user)
    {
        var search = await SearchAsync(r);
        var f = ListFilter(r);
        McoGstFilter printFilter;
        var invoiceNo = search.InvoiceNo;
        if (r.Print)
        {
            if (invoiceNo.Length == 0)
            {
                await using var tx = await sessions.BeginAsync();
                if ((f.Mode == 0 && r.PrintOption != 1) || f.Mode == 1) invoiceNo = await numbers.NextInvoiceNoAsync(tx);
                foreach (var x in search.Rows.Where(x => T(x, "MBMCInvoiceNo").Length == 0))
                    await invoices.StampAsync(tx, T(x, "MBMCNumber"), T(x, "MBMCCoverID"), invoiceNo, user);
                await tx.CommitAsync();
            }
            printFilter = new McoGstFilter(2, null, null, null, null, invoiceNo, null, null, null);
        }
        else
        {
            if (search.MultipleInvoices) throw new BusinessRuleException(McoListService.MultiInvoice);
            // PrintPreview: the criteria and the payor only (a group company / subsidiary is not applied).
            printFilter = f with { GroupCode = null, Subsidiary = null, Payor = f.Mode == 2 ? null : Left(r.Payor, 2) };
        }
        await using var s = await sessions.OpenAsync();
        var summary = await repository.SummaryAsync(s, printFilter);
        var details = await repository.DetailsAsync(s, printFilter);
        var payor = summary.Count > 0 ? T(summary[0], "PAYCode") : "";
        var company = await invoices.CompanyAsync(s, payor, r.Version != 1);
        string C(string k) => Convert.ToString(company?.GetValueOrDefault(k), CultureInfo.InvariantCulture)?.Trim() ?? "";
        var lines = summary.Select((x, i) => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["No"] = (i + 1).ToString(CultureInfo.InvariantCulture), ["BordxDate"] = x.GetValueOrDefault("MBMBordxDate"),
            ["Type"] = T(x, "INSCode") switch { "I" => "I - INDIVIDUAL", "F" => "F - FAMILY", "G" => "G - GROUP INDIVIDUAL", "H" => "H - GROUP FAMILY", _ => "" },
            ["AdultChild"] = T(x, "AdultChild") == "A" ? "Adult" : T(x, "AdultChild") == "C" ? "Child" : "", ["Principals"] = x.GetValueOrDefault("EmpCnt"),
            ["Covered"] = x.GetValueOrDefault("DepCnt"), ["Mco"] = x.GetValueOrDefault("BasicMCO"), ["Gst"] = x.GetValueOrDefault("GST") ?? 0m,
        }).ToList();
        var shown = summary.Select(x => T(x, "MBMCInvoiceNo")).FirstOrDefault(x => x.Length > 0) ?? invoiceNo;
        return exporter.Export(new ReportDocument("MCO INVOICE (GST)", SummaryColumns, lines)
        {
            Subtitles =
            [
                C("CompanyName"), C("Add1"), C("Add3"), $"ATTENTION:  {C("Contact")}", $"TEL :  {C("TelNo")}   FAX :  {C("FaxNo")}", $"Invoice No: {shown}",
            ],
            FileName = shown.Length > 0 ? shown : payor,
            ExtraSheets = [new ReportDocument("DETAILS", DetailColumns, details.Cast<IReadOnlyDictionary<string, object?>>().ToList())],
        }, ReportFileType.Excel);
    }

    private static readonly IReadOnlyList<ReportColumn> SummaryColumns =
    [
        new("No", "No"), new("BordxDate", "Bordx Date", ReportFormat.Date), new("Type", "Type"), new("AdultChild", "Adult / Child"),
        new("Principals", "Employee", ReportFormat.Text, true), new("Covered", "Dependant", ReportFormat.Text, true), new("Mco", "Basic MCO", ReportFormat.Money, true),
        new("Gst", "GST", ReportFormat.Money, true),
    ];

    private static readonly IReadOnlyList<ReportColumn> DetailColumns =
    [
        new("PAYCode", "Payor Code"), new("MBMCNumber", "MBM Number"), new("MBMCCoverID", "Cover ID"), new("MBMCRELCode", "Relationship"),
        new("MBMBordxDate", "Bordx Date", ReportFormat.Date), new("MBMPolicyNo", "Policy Number"), new("HLTCode", "Health Type"), new("MBMBatchNo", "Batch No"),
        new("MBMName", "Name"), new("MBMGroupCompany", "Group Company"), new("MBMCInvoiceNo", "Invoice Number"), new("BasicMCO", "Basic MCO", ReportFormat.Money, true),
        new("GST", "GST", ReportFormat.Money, true),
    ];

    /// <summary>cmdINVSearch_Click's checks (as MCO Invoice).</summary>
    private static void Check(McoGstRequest r)
    {
        switch (r.Criteria)
        {
            case 0:
                if (ClinicInvoiceRules.Date(r.BordxDate) is null) throw new BusinessRuleException(McoInvoiceService.NoBordx);
                if (!int.TryParse(r.BatchNo?.Trim(), out _)) throw new BusinessRuleException(McoInvoiceService.NoBatch);
                break;
            case 1:
                if (ClinicInvoiceRules.Date(r.From) is null || ClinicInvoiceRules.Date(r.To) is null) throw new BusinessRuleException(McoInvoiceService.NoRange);
                break;
            case 2:
                if (string.IsNullOrWhiteSpace(r.InvoiceNo)) throw new BusinessRuleException(McoInvoiceService.NoInvoice);
                break;
            default:
                throw new BusinessRuleException(McoInvoiceService.NoBordx);
        }
        if (r.Criteria == 2) return;
        if ((r.Version ?? 0) == 0 && string.IsNullOrWhiteSpace(r.Payor)) throw new BusinessRuleException(McoInvoiceService.NoPayor);
        if (r.Version == 1 && string.IsNullOrWhiteSpace(r.GroupCompany)) throw new BusinessRuleException(McoInvoiceService.NoCompany);
    }

    /// <summary>SelectionCriterial (GST): group / subsidiary whenever given, the payor for the payor version.</summary>
    private static McoGstFilter ListFilter(McoGstRequest r) =>
        new(r.Criteria!.Value, ClinicInvoiceRules.Date(r.BordxDate), int.TryParse(r.BatchNo?.Trim(), out var b) ? b : null, ClinicInvoiceRules.Date(r.From),
            ClinicInvoiceRules.Date(r.To), r.InvoiceNo?.Trim(), Left(r.GroupCompany, 4), string.IsNullOrWhiteSpace(r.Subsidiary) ? null : r.Subsidiary.Trim(),
            (r.Version ?? 0) == 0 ? Left(r.Payor, 2) : null);

    private static string? Left(string? v, int n) => string.IsNullOrWhiteSpace(v) ? null : v.Trim().Length <= n ? v.Trim() : v.Trim()[..n].Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";
}
