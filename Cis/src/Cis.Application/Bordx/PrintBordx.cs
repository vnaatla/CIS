using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Bordx;

/// <summary>Print Bordx (frmPrintBordxExcel): Print Details / Print Summary.</summary>
public sealed record PrintBordxRequest
{
#pragma warning disable CS1591
    public string BordxNo { get; init; } = "";
    /// <summary>MediExpress, Health Connect or Sihat Johor.</summary>
    public string? Company { get; init; }
    /// <summary>0 Invoice (the TAX templates), 1 Disbursement Note (the "_20" templates).</summary>
    public int? BillType { get; init; }
    /// <summary>0 by group company, 1 by department, 2 by relationship, 3 by cost center, 4 by cost center-branch, 5 by subsidiary.</summary>
    public int ReportType { get; init; }
    public string? Month { get; init; }
    public string? Year { get; init; }
    /// <summary>Print (stamps the bordereaux as printed) instead of Preview.</summary>
    public bool Print { get; init; }
    /// <summary>Summary only (Print Summary).</summary>
    public bool Summary { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class PrintBordxRequestValidator : Common.RequestValidator<PrintBordxRequest>
{
    /// <summary>Defines the rules.</summary>
    public PrintBordxRequestValidator()
    {
        RuleFor(r => r.BordxNo).MaximumLength(20);
        RuleFor(r => r.Month).MaximumLength(20);
        RuleFor(r => r.Year).MaximumLength(4);
    }
}

/// <summary>GetBordxInfo + PrintBordxListing.</summary>
public record PrintBordxInfo(Dictionary<string, object?> Bordx, bool Printed, List<Dictionary<string, object?>> Rows, decimal TotalAmount);

/// <summary>Data access.</summary>
public interface IPrintBordxRepository
{
#pragma warning disable CS1591
    Task<Dictionary<string, object?>?> InfoAsync(IDbSession session, string bordxNo);
    Task<List<Dictionary<string, object?>>> RowsAsync(IDbSession session, string bordxNo);
    Task<List<Dictionary<string, object?>>> SummaryAsync(IDbSession session, string bordxNo);
    Task UpdateAsync(IDbSession session, string bordxNo, string month, string year, string user);
#pragma warning restore CS1591
}

/// <summary>Claim Bordx > Print Bordx.</summary>
public interface IPrintBordxService
{
    /// <summary>cmdSearch_Click (with GetBordxInfo).</summary>
    Task<PrintBordxInfo> FindAsync(string? bordxNo);

    /// <summary>cmdPrintDetails_Click / CmdPrintSummary_Click and their previews: the pay advice (Excel).</summary>
    Task<FileOutput> PrintAsync(PrintBordxRequest request, string user);
}

/// <inheritdoc />
public sealed class PrintBordxService(IDbSessions sessions, IPrintBordxRepository repository, IReportExporter exporter) : IPrintBordxService
{
    /// <summary>Desktop texts.</summary>
    public const string NoBordx = "Please enter Bordx Reference No", NoRecord = "No record found ", NoCompany = "Please Select company name",
        NoBillType = "Please Select Bill Type", NoMonth = "Please keyin Month or Year !", WrongCompany = "Please Select Health Connect Or Sihat Johor..!";

    /// <summary>The companies of cboCompany.</summary>
    public static readonly string[] Companies = ["MediExpress", "Health Connect", "Sihat Johor"];

    /// <inheritdoc />
    public async Task<PrintBordxInfo> FindAsync(string? bordxNo)
    {
        if (string.IsNullOrWhiteSpace(bordxNo)) throw new BusinessRuleException(NoBordx);
        await using var s = await sessions.OpenAsync();
        var no = bordxNo.Trim();
        var rows = await repository.RowsAsync(s, no);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var info = await repository.InfoAsync(s, no) ?? new Dictionary<string, object?>();
        foreach (var x in rows) x["Name"] = T(x, "CoverID") == "00" ? T(x, "MBMName") : T(x, "MBMCName");
        return new PrintBordxInfo(info, Printed(info), rows, rows.Sum(x => N(x, "CLMBordxAmt")));
    }

    /// <inheritdoc />
    public async Task<FileOutput> PrintAsync(PrintBordxRequest r, string user)
    {
        if (string.IsNullOrWhiteSpace(r.BordxNo)) throw new BusinessRuleException(NoBordx);
        if (!r.Summary)
        {
            if (string.IsNullOrWhiteSpace(r.Company)) throw new BusinessRuleException(NoCompany);
            if (r.BillType is not (0 or 1)) throw new BusinessRuleException(NoBillType);
            if (!Companies.Contains(r.Company.Trim())) throw new BusinessRuleException(WrongCompany);
        }
        var no = r.BordxNo.Trim();
        await using var s = await sessions.OpenAsync();
        var info = await repository.InfoAsync(s, no) ?? throw new NotFoundException(NoRecord);
        var printed = Printed(info);
        // VerifyPrint: a bordereaux not printed yet needs the month and year.
        if (!printed && (string.IsNullOrWhiteSpace(r.Month) || string.IsNullOrWhiteSpace(r.Year))) throw new BusinessRuleException(NoMonth);
        var month = printed ? T(info, "BORDXPrintMonth") : r.Month!.Trim();
        var year = printed ? T(info, "BORDXPrintYear") : r.Year!.Trim();
        var rows = await repository.RowsAsync(s, no);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var summary = await repository.SummaryAsync(s, no);
        foreach (var x in rows)
        {
            x["Name"] = T(x, "CoverID") == "00" ? T(x, "MBMName") : T(x, "MBMCName");
            x["Group"] = r.ReportType switch
            {
                1 => T(x, "MBMDepartment"), 2 => T(x, "Relation") is { Length: > 0 } rel ? rel : T(x, "MBMCRELCode"), 3 => T(x, "MBMCostCenter"),
                4 => $"{T(x, "MBMCostCenter")} / {T(x, "MBMBranch")}", 5 => T(x, "Subsidiary"), _ => T(x, "MBMGroupCompany"),
            };
        }
        if (r.Print && !printed)
        {
            await using var tx = await sessions.BeginAsync();
            await repository.UpdateAsync(tx, no, month, year, user);
            await tx.CommitAsync();
        }
        var head = summary.FirstOrDefault();
        var credit = no.StartsWith("KCN", StringComparison.OrdinalIgnoreCase);
        var title = (credit ? "CREDIT NOTE - " : "") + (r.Company is "MediExpress" ? "MEDIEXPRESS " : "") + (r.BillType == 0 ? "INVOICE" : "DISBURSEMENT NOTE");
        var subtitles = new List<string>
        {
            T(head, "subname") is { Length: > 0 } sub ? sub : T(head, "MBMGroupCompany"), Join(head, "GRPCoAdd1", "GRPCoAdd2", "GRPCoAdd3", "GRPPostCode"),
            $"Attn: {(T(head, "PIC") is { Length: > 0 } pic ? pic : T(head, "GRPCoContact"))}",
            $"Bordx No: {no}   Bordx Date: {(info.GetValueOrDefault("BordxDate") as DateTime?)?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}",
            $"Actual Outpatient medical cost incurred for the month of {month} {year}", $"Report: {ReportTypes[Math.Clamp(r.ReportType, 0, 5)]}",
        };
        var groups = rows.GroupBy(x => T(x, "Group")).OrderBy(g => g.Key, StringComparer.Ordinal)
            .Select(g => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>
            {
                ["Group"] = g.Key, ["Cases"] = g.Count(), ["Actual"] = g.Sum(x => N(x, "CLMTotalActual")), ["CoPay"] = g.Sum(x => N(x, "CLMCoPayCLNColl")),
                ["Approved"] = g.Sum(x => N(x, "CLMFinalApprove")), ["Mco"] = g.Sum(x => N(x, "CLMMCOProcFee")), ["Bordx"] = g.Sum(x => N(x, "CLMBordxAmt")),
            }).ToList();
        var details = rows.OrderBy(x => T(x, "Group"), StringComparer.Ordinal).ThenBy(x => T(x, "CASNumber"), StringComparer.Ordinal)
            .Cast<IReadOnlyDictionary<string, object?>>().ToList();
        return exporter.Export(new ReportDocument(title, SummaryColumns, groups)
        {
            Subtitles = subtitles, FileName = $"PrintBordx_{no}",
            ExtraSheets = r.Summary ? [] : [new ReportDocument("DETAILS", DetailColumns, details)],
        }, ReportFileType.Excel);
    }

    private static readonly string[] ReportTypes = ["By Group Company", "By Department", "By Relationship", "By Cost Center", "By Cost Center - Branch", "By Subsidiary"];

    private static readonly IReadOnlyList<ReportColumn> SummaryColumns =
    [
        new("Group", "Group"), new("Cases", "Cases", ReportFormat.Text, true), new("Actual", "Bill Amount", ReportFormat.Money, true),
        new("CoPay", "Co-Payment", ReportFormat.Money, true), new("Approved", "Approved", ReportFormat.Money, true), new("Mco", "MCO Fee", ReportFormat.Money, true),
        new("Bordx", "Bordx Amount", ReportFormat.Money, true),
    ];

    private static readonly IReadOnlyList<ReportColumn> DetailColumns =
    [
        new("Group", "Group"), new("CASNumber", "Case No"), new("MBMnumber", "Member No"), new("CoverID", "Cov"), new("Name", "Patient"), new("MBMEmployeeNo", "Employee No"),
        new("MBMDepartment", "Department"), new("DateVisit", "Visit Date", ReportFormat.Date), new("CLNName", "Clinic"), new("CLNType", "Clinic Type"),
        new("ReasonDesc1", "Diagnosis"), new("MCDays", "MC Days"), new("INVPayTo", "Pay To"), new("CLMTotalActual", "Bill Amount", ReportFormat.Money, true),
        new("CLMCoPayCLNColl", "Co-Payment", ReportFormat.Money, true), new("CLMFinalApprove", "Approved", ReportFormat.Money, true),
        new("CLMMCOProcFee", "MCO Fee", ReportFormat.Money, true), new("CLMBordxAmt", "Bordx Amount", ReportFormat.Money, true),
    ];

    private static bool Printed(IReadOnlyDictionary<string, object?> info) => T(info, "BordxPrintedStatus") is "1" or "True";

    private static string Join(IReadOnlyDictionary<string, object?>? r, params string[] keys) => string.Join(", ", keys.Select(k => T(r, k)).Where(v => v.Length > 0));

    private static string T(IReadOnlyDictionary<string, object?>? r, string k) =>
        r is not null && r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
