using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Mco;

/// <summary>MCO Invoice search (frmMCOInvoiceNew).</summary>
public sealed record McoInvoiceRequest
{
    /// <summary>0 Bordereaux Batch, 1 Bordereaux Range, 2 Invoice No.</summary>
    public int? Criteria { get; init; }
    /// <summary>dd/mm/yyyy.</summary>
    public string? BordxDate { get; init; }
    /// <summary>Batch no.</summary>
    public string? BatchNo { get; init; }
    /// <summary>Range from.</summary>
    public string? From { get; init; }
    /// <summary>Range to.</summary>
    public string? To { get; init; }
    /// <summary>Invoice no.</summary>
    public string? InvoiceNo { get; init; }
    /// <summary>0 Invoice - by Payor, 1 by Group Company, 2 by Department.</summary>
    public int? Version { get; init; }
    /// <summary>"code - name".</summary>
    public string? Payor { get; init; }
    /// <summary>"code - name".</summary>
    public string? GroupCompany { get; init; }
    /// <summary>Department.</summary>
    public string? Department { get; init; }
    /// <summary>0 Print New, 1 Reprint.</summary>
    public int? PrintOption { get; init; }
    /// <summary>Print (number and stamp) instead of Print Preview.</summary>
    public bool Print { get; init; }
}

/// <summary>Lengths.</summary>
public sealed class McoInvoiceRequestValidator : Common.RequestValidator<McoInvoiceRequest>
{
    /// <summary>Defines the rules.</summary>
    public McoInvoiceRequestValidator()
    {
        RuleFor(r => r.InvoiceNo).MaximumLength(30);
        RuleFor(r => r.Payor).MaximumLength(200);
        RuleFor(r => r.GroupCompany).MaximumLength(300);
        RuleFor(r => r.Department).MaximumLength(200);
    }
}

/// <summary>A listed member or covered person.</summary>
public record McoInvoiceRow(string Payor, string Member, string Cover, string Type, DateTime? BordxDate, string PolicyNo, string BatchNo, string Name,
    string GroupCompany, bool Invoiced, string InvoiceNo);

/// <summary>The list, its counts and the invoice no. shown (the last one found).</summary>
public record McoInvoiceResult(List<McoInvoiceRow> Rows, int Principals, int Supplementaries, string InvoiceNo, bool MultipleInvoices);

/// <summary>cis.McoInv_* parameters.</summary>
public record McoInvoiceFilter(int Mode, DateTime? BordxDate, int? BatchNo, DateTime? From, DateTime? To, string? InvoiceNo, int Version, string? Payor,
    string? GroupCode, string? Department, bool Reprint);

/// <summary>Data access.</summary>
public interface IMcoInvoiceRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> PrincipalsAsync(IDbSession session, McoInvoiceFilter filter);
    Task<List<Dictionary<string, object?>>> SuppsAsync(IDbSession session, McoInvoiceFilter filter, string? member);
    Task<List<Dictionary<string, object?>>> SummaryAsync(IDbSession session, McoInvoiceFilter filter);
    Task<Dictionary<string, object?>?> CompanyAsync(IDbSession session, string code, bool byPayor);
    Task StampAsync(IDbSession session, string member, string cover, string invoiceNo, string user);
    Task<List<string>> DepartmentsAsync(IDbSession session, string company);
#pragma warning restore CS1591
}

/// <summary>MCO Fees > MCO Invoice (frmMCOInvoiceNew).</summary>
public interface IMcoInvoiceService
{
    /// <summary>cboINVGrpCom_LostFocus: the company's departments.</summary>
    Task<List<string>> DepartmentsAsync(string? groupCompany);

    /// <summary>cmdINVSearch_Click.</summary>
    Task<McoInvoiceResult> SearchAsync(McoInvoiceRequest request);

    /// <summary>cmdINVPrint_Click / cmdINVPrintPreview_Click: the invoice (Excel).</summary>
    Task<FileOutput> PrintAsync(McoInvoiceRequest request, string user);
}

/// <inheritdoc />
public sealed class McoInvoiceService(IDbSessions sessions, IMcoInvoiceRepository repository, IMcoListRepository numbers, IReportExporter exporter,
    TimeProvider clock) : IMcoInvoiceService
{
    /// <summary>Desktop texts.</summary>
    public const string NoBordx = "Please Enter Bordereaux Date!", NoBatch = "Please Enter Batch Number!", NoRange = "Please enter valid Bordereaux Range !",
        NoInvoice = "Please enter Invoice Number!", NoPayor = "Please Select a Payor!", NoCompany = "Please Select a Group Company!",
        NoDepartment = "Please Select a Department!", NoRecord = "No record found!";

    /// <inheritdoc />
    public async Task<List<string>> DepartmentsAsync(string? groupCompany)
    {
        if (string.IsNullOrWhiteSpace(groupCompany)) return [];
        var g = groupCompany.Trim();
        // The desktop takes the company name from position 7 of "CODE - NAME" up to the next "-".
        var name = g.Length > 6 ? g[6..] : g;
        if (name.IndexOf('-') is > 0 and var i) name = name[..i];
        await using var s = await sessions.OpenAsync();
        return await repository.DepartmentsAsync(s, name.Trim());
    }

    /// <inheritdoc />
    public async Task<McoInvoiceResult> SearchAsync(McoInvoiceRequest r)
    {
        var f = Filter(r);
        await using var s = await sessions.OpenAsync();
        var principals = await repository.PrincipalsAsync(s, f);
        var rows = new List<McoInvoiceRow>();
        var listed = new HashSet<string>(StringComparer.Ordinal);
        var invoices = new List<string>();
        int supps = 0;
        foreach (var p in principals)
        {
            var inv = T(p, "MBMCInvoiceNo");
            if (inv.Length > 0) invoices.Add(inv);
            rows.Add(new McoInvoiceRow(T(p, "PAYCode"), T(p, "MBMNumber"), "00", T(p, "MBMMemberType"), D(p, "MBMBordxDate"), T(p, "MBMPolicyNo"), T(p, "MBMBatchNo"),
                T(p, "MBMName"), T(p, "MBMGroupCompany"), T(p, "MBMCInvoiced") == "Y", inv));
            if (T(p, "INSCode") is "F" or "H" && listed.Add(T(p, "MBMNumber")))
                supps += AddSupps(rows, await repository.SuppsAsync(s, f, T(p, "MBMNumber")), null, invoices);
        }
        supps += AddSupps(rows, await repository.SuppsAsync(s, f, null), listed, invoices);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        return new McoInvoiceResult(rows, principals.Count, supps, invoices.Count > 0 ? invoices[^1] : "", invoices.Distinct().Count() > 1);
    }

    /// <inheritdoc />
    public async Task<FileOutput> PrintAsync(McoInvoiceRequest r, string user)
    {
        var search = await SearchAsync(r);
        if (search.MultipleInvoices) throw new BusinessRuleException(McoListService.MultiInvoice);
        var f = Filter(r);
        await using var s = await sessions.OpenAsync();
        var lines = await repository.SummaryAsync(s, f);
        if (lines.Count == 0) throw new NotFoundException(NoRecord);
        var byPayor = f.Version == 0;
        var code = byPayor ? T(lines[0], "PAYCode") : T(lines[0], "GRPCoCode");
        var company = await repository.CompanyAsync(s, code, byPayor);
        var invoiceNo = T(lines[0], "MBMCInvoiceNo");
        var invoiceDate = D(lines[0], "MBMCInvoiceDate") ?? clock.GetLocalNow().DateTime.Date;
        if (r.Print && search.InvoiceNo.Length == 0)
        {
            await using var tx = await sessions.BeginAsync();
            if ((f.Mode == 0 && !f.Reprint) || f.Mode == 1) invoiceNo = await numbers.NextInvoiceNoAsync(tx);
            foreach (var x in search.Rows.Where(x => !x.Invoiced)) await repository.StampAsync(tx, x.Member, x.Cover, invoiceNo, user);
            await tx.CommitAsync();
            invoiceDate = clock.GetLocalNow().DateTime.Date;
        }
        var summary = Summary(lines);
        string C(string k) => Convert.ToString(company?.GetValueOrDefault(k), CultureInfo.InvariantCulture)?.Trim() ?? "";
        return exporter.Export(new ReportDocument("MCO INVOICE", SummaryColumns, summary)
        {
            Subtitles =
            [
                .. f.Version == 2 ? new[] { T(lines[0], "MBMDepartment") } : [], C("CompanyName"), C("Add1"), C("Add2"), C("Add3"), $"ATTENTION:  {C("Contact")}",
                $"TEL :  {C("TelNo")}   FAX :  {C("FaxNo")}", $"Invoice No: {invoiceNo}   Date: {invoiceDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}",
            ],
            FileName = invoiceNo.Length > 0 ? $"MCOInvoice_{invoiceNo}" : "MCOInvoice",
        }, ReportFileType.Excel);
    }

    private static readonly IReadOnlyList<ReportColumn> SummaryColumns =
    [
        new("No", "No"), new("BordxDate", "Bordx Date", ReportFormat.Date), new("Type", "Type"), new("AdultChild", "Adult / Child"),
        new("Principals", "Principal", ReportFormat.Text, true), new("Covered", "Covered", ReportFormat.Text, true), new("Mco", "MCO Fee", ReportFormat.Money, true),
    ];

    /// <summary>The invoice lines: one per insurer / bordx date / adult-child, principals (member type other than S), all covered, MCO fee.</summary>
    internal static List<IReadOnlyDictionary<string, object?>> Summary(List<Dictionary<string, object?>> lines)
    {
        var result = new List<IReadOnlyDictionary<string, object?>>();
        foreach (var g in lines.GroupBy(x => (T(x, "INSCode"), D(x, "MBMBordxDate"), T(x, "MBMAdultChild"))))
        {
            decimal Q(Dictionary<string, object?> x) => Num(x, "MBMQty");
            var ins = g.Key.Item1 switch { "I" => "I - INDIVIDUAL", "F" => "F - FAMILY", "G" => "G - GROUP INDIVIDUAL", "H" => "H - GROUP FAMILY", _ => "" };
            result.Add(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            {
                ["No"] = (result.Count + 1).ToString(CultureInfo.InvariantCulture), ["BordxDate"] = g.Key.Item2, ["Type"] = ins,
                ["AdultChild"] = g.Key.Item3 == "A" ? "Adult" : g.Key.Item3 == "C" ? "Child" : "",
                ["Principals"] = g.Where(x => T(x, "MBMMemberType") != "S").Sum(Q), ["Covered"] = g.Sum(Q), ["Mco"] = g.Sum(x => Num(x, "MCOFee")),
            });
        }
        return result;
    }

    private static int AddSupps(List<McoInvoiceRow> rows, List<Dictionary<string, object?>> supps, HashSet<string>? skip, List<string> invoices)
    {
        var n = 0;
        foreach (var c in supps)
        {
            if (skip is not null && skip.Contains(T(c, "MBMCNumber"))) continue;
            var inv = T(c, "MBMCInvoiceNo");
            if (inv.Length > 0) invoices.Add(inv);
            rows.Add(new McoInvoiceRow(T(c, "PAYCode"), T(c, "MBMCNumber"), T(c, "MBMCCoverID"), T(c, "MBMCRELCode"), D(c, "MBMCBordxDate"), T(c, "MBMPolicyNo"),
                T(c, "MBMCBatchNo"), T(c, "MBMCName"), T(c, "MBMGroupCompany"), inv.Length > 0, inv));
            n++;
        }
        return n;
    }

    /// <summary>cmdINVSearch_Click's checks and SelectionCriterial.</summary>
    private static McoInvoiceFilter Filter(McoInvoiceRequest r)
    {
        DateTime? bordx = null, from = null, to = null;
        int? batch = null;
        switch (r.Criteria)
        {
            case 0:
                bordx = ClinicInvoiceRules.Date(r.BordxDate) ?? throw new BusinessRuleException(NoBordx);
                if (string.IsNullOrWhiteSpace(r.BatchNo)) throw new BusinessRuleException(NoBatch);
                batch = int.TryParse(r.BatchNo.Trim(), out var b) ? b : throw new BusinessRuleException(NoBatch);
                break;
            case 1:
                from = ClinicInvoiceRules.Date(r.From);
                to = ClinicInvoiceRules.Date(r.To);
                if (from is null || to is null) throw new BusinessRuleException(NoRange);
                break;
            case 2:
                if (string.IsNullOrWhiteSpace(r.InvoiceNo)) throw new BusinessRuleException(NoInvoice);
                break;
            default:
                throw new BusinessRuleException(NoBordx);
        }
        var version = r.Version ?? 0;
        if (r.Criteria != 2)
        {
            if (version == 0 && string.IsNullOrWhiteSpace(r.Payor)) throw new BusinessRuleException(NoPayor);
            if (version is 1 or 2 && string.IsNullOrWhiteSpace(r.GroupCompany)) throw new BusinessRuleException(NoCompany);
        }
        if (version == 2 && string.IsNullOrWhiteSpace(r.Department)) throw new BusinessRuleException(NoDepartment);
        return new McoInvoiceFilter(r.Criteria!.Value, bordx, batch, from, to, r.InvoiceNo?.Trim(), version, Left(r.Payor, 2), Left(r.GroupCompany, 4),
            r.Department?.Trim(), r.PrintOption == 1);
    }

    private static string? Left(string? v, int n) => string.IsNullOrWhiteSpace(v) ? null : v.Trim().Length <= n ? v.Trim() : v.Trim()[..n].Trim();

    private static string T(IReadOnlyDictionary<string, object?>? r, string k) =>
        r is not null && r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static DateTime? D(IReadOnlyDictionary<string, object?> r, string k) => r.GetValueOrDefault(k) as DateTime?;

    private static decimal Num(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0;
}
