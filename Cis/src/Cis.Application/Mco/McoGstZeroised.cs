using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Mco;

/// <summary>GST Zeroised search / print (frmMCOFeeGSTZeroised): the GST filters plus the policy no.</summary>
public sealed record McoZeroRequest
{
#pragma warning disable CS1591
    public McoGstRequest Search { get; init; } = new();
    public string? PolicyNo { get; init; }
    /// <summary>Print: the checked members (member-cover) to number and stamp.</summary>
    public List<string> Checked { get; init; } = [];
    /// <summary>Print Excel: by subsidiary (the Yes answer).</summary>
    public bool BySubsidiary { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class McoZeroRequestValidator : Common.RequestValidator<McoZeroRequest>
{
    /// <summary>Defines the rules.</summary>
    public McoZeroRequestValidator()
    {
        RuleFor(r => r.PolicyNo).MaximumLength(50);
        RuleFor(r => r.Checked).Must(c => c.Count <= 20000);
    }
}

/// <summary>Data access.</summary>
public interface IMcoZeroRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> ListAsync(IDbSession session, McoGstFilter filter, string? policyNo);
    Task<Dictionary<string, object?>?> CompanyAsync(IDbSession session, string? subsidiary, string? subsidiary2, string? group);
#pragma warning restore CS1591
}

/// <summary>MCO Fees > MCO Invoice GST Zeroised (frmMCOFeeGSTZeroised).</summary>
public interface IMcoZeroService
{
    /// <summary>cmdINVSearch_Click.</summary>
    Task<McoGstResult> SearchAsync(McoZeroRequest request);

    /// <summary>cmdINVPrint_Click: numbers and stamps the checked members, then prints the zeroised invoice.</summary>
    Task<FileOutput> PrintAsync(McoZeroRequest request, string user);

    /// <summary>CmdPrintExcel_Click: the zeroised invoice of the searched invoice no.</summary>
    Task<FileOutput> PrintInvoiceAsync(McoZeroRequest request);
}

/// <inheritdoc />
public sealed class McoZeroService(IDbSessions sessions, IMcoZeroRepository repository, IMcoInvoiceRepository invoices, IMcoListRepository numbers,
    IReportExporter exporter, TimeProvider clock) : IMcoZeroService
{
    /// <summary>Desktop texts.</summary>
    public const string Invoiced = "The selected member already invoice generated,Please check..!", NoInvoiceSearch = "Please search with invoice no.",
        NoChecked = "Please select a record.";

    /// <inheritdoc />
    public async Task<McoGstResult> SearchAsync(McoZeroRequest r)
    {
        Check(r.Search);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.ListAsync(s, Filter(r.Search), Blank(r.PolicyNo));
        if (rows.Count == 0) throw new NotFoundException(McoInvoiceService.NoRecord);
        var inv = rows.Select(x => T(x, "MBMCInvoiceNo")).Where(x => x.Length > 0).ToList();
        var princ = rows.Count(x => T(x, "RelCode") == "P");
        return new McoGstResult(rows, princ, rows.Count - princ, inv.Count > 0 ? inv[^1] : "", inv.Distinct().Count() > 1);
    }

    /// <inheritdoc />
    public async Task<FileOutput> PrintAsync(McoZeroRequest r, string user)
    {
        var search = await SearchAsync(r);
        var invoiceNo = search.InvoiceNo;
        if (invoiceNo.Length == 0)
        {
            var picked = search.Rows.Where(x => r.Checked.Contains($"{T(x, "MBMCNumber")}-{T(x, "MBMCCoverID")}")).ToList();
            if (picked.Count == 0) throw new BusinessRuleException(NoChecked);
            if (picked.Any(x => T(x, "MBMCInvoiceNo").Length > 0)) throw new BusinessRuleException(Invoiced);
            var f = Filter(r.Search);
            await using var tx = await sessions.BeginAsync();
            if ((f.Mode == 0 && r.Search.PrintOption != 1) || f.Mode == 1) invoiceNo = await numbers.NextInvoiceNoAsync(tx);
            foreach (var x in picked) await invoices.StampAsync(tx, T(x, "MBMCNumber"), T(x, "MBMCCoverID"), invoiceNo, user);
            await tx.CommitAsync();
        }
        if (invoiceNo.Length == 0) throw new BusinessRuleException(McoInvoiceService.NoInvoice);
        return await Zeroised(invoiceNo, Blank(r.Search.Subsidiary), false);
    }

    /// <inheritdoc />
    public async Task<FileOutput> PrintInvoiceAsync(McoZeroRequest r)
    {
        if (r.Search.Criteria != 2 || string.IsNullOrWhiteSpace(r.Search.InvoiceNo)) throw new BusinessRuleException(NoInvoiceSearch);
        return await Zeroised(r.Search.InvoiceNo.Trim(), Blank(r.Search.Subsidiary), r.BySubsidiary);
    }

    /// <summary>PrintZeroisedInvoice: the member rows (sheet 2 in the desktop) and the invoice block (count, TPA fee total).</summary>
    private async Task<FileOutput> Zeroised(string invoiceNo, string? subsidiary, bool bySubsidiary)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.ListAsync(s, new McoGstFilter(2, null, null, null, null, invoiceNo, null, null, null), null);
        var first = rows.FirstOrDefault();
        var sub2 = bySubsidiary ? Blank(T(first, "Subsidiary")) : null;
        var company = await repository.CompanyAsync(s, subsidiary, sub2, subsidiary is null && sub2 is null ? T(first, "MBMGroupCompany") : null);
        string C(string k) => Convert.ToString(company?.GetValueOrDefault(k), CultureInfo.InvariantCulture)?.Trim() is { Length: > 0 } v ? v : "NA";
        var total = rows.Sum(x => x.GetValueOrDefault("TPAFEE") is { } v and not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : 0);
        var name = C("Name");
        var invoice = new ReportDocument("TAX INVOICE (GST ZEROISED)",
            [new ReportColumn("Item", "Description"), new ReportColumn("Qty", "Persons"), new ReportColumn("Amount", "Amount (RM)", ReportFormat.Money)],
            [new Dictionary<string, object?> { ["Item"] = "TPA fee", ["Qty"] = rows.Count.ToString(CultureInfo.InvariantCulture), ["Amount"] = total }])
        {
            Subtitles =
            [
                $"Attn: {C("Contact")}", name, C("Add1"), C("Add2"), C("Add3"),
                $"Invoice No: {invoiceNo}   Date: {clock.GetLocalNow().ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}",
            ],
            FileName = $"{invoiceNo}-{name.Replace("'", "").Replace(".", "")}",
            ExtraSheets = [new ReportDocument("MEMBERS", MemberColumns, rows.Cast<IReadOnlyDictionary<string, object?>>().ToList())],
        };
        return exporter.Export(invoice, ReportFileType.Excel);
    }

    private static readonly IReadOnlyList<ReportColumn> MemberColumns =
    [
        new("MBMCNumber", "Member No"), new("MBMCostCenter", "Cost Center"), new("MBMDepartment", "Department"), new("MBMDivision", "Division"),
        new("MBMEmployeeNo", "Employee No"), new("MBMName", "Name"), new("Priname", "Principal"), new("IC", "IC"), new("DOB", "DOB", ReportFormat.Date),
        new("SEX", "Sex"), new("MBMCCoverID", "Cover"), new("Relation", "Relation"), new("MBMBordxDate", "Bordx Date", ReportFormat.Date),
        new("EFFDATE", "Eff Date", ReportFormat.Date), new("EXPDATE", "Exp Date", ReportFormat.Date), new("MBMPolicyNo", "Policy No"),
        new("MBMGroupCompany", "Group Company"), new("Subsidiary", "Subsidiary"), new("MBMBranch", "Branch"), new("policystatus", "Policy Status"),
        new("clinicalremark", "Clinical Remark"), new("memberremark", "Member Remark"), new("DayInsuredPlus1", "Days Insured"),
        new("BasicMCO", "Basic MCO", ReportFormat.Money), new("TPAFEE", "TPA Fee", ReportFormat.Money, true),
    ];

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
                return;
            default:
                throw new BusinessRuleException(McoInvoiceService.NoBordx);
        }
        if ((r.Version ?? 0) == 0 && string.IsNullOrWhiteSpace(r.Payor)) throw new BusinessRuleException(McoInvoiceService.NoPayor);
        if (r.Version == 1 && string.IsNullOrWhiteSpace(r.GroupCompany)) throw new BusinessRuleException(McoInvoiceService.NoCompany);
    }

    private static McoGstFilter Filter(McoGstRequest r) =>
        new(r.Criteria!.Value, ClinicInvoiceRules.Date(r.BordxDate), int.TryParse(r.BatchNo?.Trim(), out var b) ? b : null, ClinicInvoiceRules.Date(r.From),
            ClinicInvoiceRules.Date(r.To), r.InvoiceNo?.Trim(), Left(r.GroupCompany, 4), Blank(r.Subsidiary), (r.Version ?? 0) == 0 ? Left(r.Payor, 2) : null);

    private static string? Left(string? v, int n) => string.IsNullOrWhiteSpace(v) ? null : v.Trim().Length <= n ? v.Trim() : v.Trim()[..n].Trim();

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?>? r, string k) =>
        r is not null && r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";
}
