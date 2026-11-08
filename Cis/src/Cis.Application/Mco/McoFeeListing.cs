using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Mco;

/// <summary>Fee Listing search (frmMCOFeeListNew).</summary>
public sealed record McoListRequest
{
    /// <summary>0 Bordereaux Batch, 1 Bordereaux Range, 2 Invoice.</summary>
    public int? Criteria { get; init; }
    /// <summary>dd/mm/yyyy.</summary>
    public string? BordxDate { get; init; }
    /// <summary>Batch no.</summary>
    public string? BatchNo { get; init; }
    /// <summary>Range from (dd/mm/yyyy).</summary>
    public string? From { get; init; }
    /// <summary>Range to (dd/mm/yyyy).</summary>
    public string? To { get; init; }
    /// <summary>Invoice no.</summary>
    public string? InvoiceNo { get; init; }
    /// <summary>0 by Payor, 1 by Group Company, 2 by Department.</summary>
    public int? Version { get; init; }
    /// <summary>"code - name" or the code.</summary>
    public string? Payor { get; init; }
    /// <summary>"code - name" or the code.</summary>
    public string? GroupCompany { get; init; }
    /// <summary>Print status: N (not invoiced) or Y.</summary>
    public string? Invoiced { get; init; }
}

/// <summary>Lengths.</summary>
public sealed class McoListRequestValidator : Common.RequestValidator<McoListRequest>
{
    /// <summary>Defines the rules.</summary>
    public McoListRequestValidator()
    {
        RuleFor(r => r.InvoiceNo).MaximumLength(30);
        RuleFor(r => r.Payor).MaximumLength(200);
        RuleFor(r => r.GroupCompany).MaximumLength(300);
    }
}

/// <summary>Print / Print Preview of the listing.</summary>
public sealed record McoListPrintRequest
{
    /// <summary>The search.</summary>
    public McoListRequest Search { get; init; } = new();
    /// <summary>0 MCO Invoice, 1 Fee Listing.</summary>
    public int? ReportType { get; init; }
    /// <summary>The selected row's payor (lvwMCO.SelectedItem).</summary>
    public string? SelectedPayor { get; init; }
    /// <summary>The selected row's batch no.</summary>
    public string? SelectedBatch { get; init; }
    /// <summary>Print (numbers and stamps a new invoice) instead of Preview.</summary>
    public bool Print { get; init; }
}

/// <summary>Lengths.</summary>
public sealed class McoListPrintRequestValidator : Common.RequestValidator<McoListPrintRequest>
{
    /// <summary>Defines the rules.</summary>
    public McoListPrintRequestValidator()
    {
        RuleFor(r => r.SelectedPayor).MaximumLength(10);
        RuleFor(r => r.SelectedBatch).MaximumLength(10);
    }
}

/// <summary>Search result: the rows, the principal / supplementary counts and what the desktop enabled.</summary>
public record McoListResult(List<Dictionary<string, object?>> Rows, int Principals, int Supplementaries, bool CanPrint, bool MultipleInvoices);

/// <summary>Data access.</summary>
public interface IMcoListRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, McoListFilter filter);
    Task<(List<(string Code, string Name)> Payors, List<(string Code, string Name)> Groups)> LookupsAsync(IDbSession session);
    Task<Dictionary<string, object?>?> CompanyAsync(IDbSession session, string code);
    Task<string> NextInvoiceNoAsync(IDbSession session);
    Task StampAsync(IDbSession session, string member, string cover, string invoiceNo, string user, decimal mco);
#pragma warning restore CS1591
}

/// <summary>cis.McoList_Search parameters.</summary>
public record McoListFilter(int Mode, DateTime? BordxDate, int? BatchNo, DateTime? From, DateTime? To, string? InvoiceNo, string Invoiced, string? Payor, string? GroupCode);

/// <summary>MCO Fees > Fee Listing (frmMCOFeeListNew).</summary>
public interface IMcoListService
{
    /// <summary>Form_Load.</summary>
    Task<(List<(string Code, string Name)> Payors, List<(string Code, string Name)> Groups)> LookupsAsync();

    /// <summary>cmdMCOSearch_Click.</summary>
    Task<McoListResult> SearchAsync(McoListRequest request);

    /// <summary>cmdMCOPrint_Click / cmdMCOPrintPreview_Click.</summary>
    Task<FileOutput> PrintAsync(McoListPrintRequest request, string user);
}

/// <inheritdoc />
public sealed class McoListService(IDbSessions sessions, IMcoListRepository repository, IReportExporter exporter, TimeProvider clock) : IMcoListService
{
    /// <summary>Desktop texts.</summary>
    public const string NoCriteria = "Please Enter MCO Criteria!", NoInvoice = "Please Enter Invoice Number!",
        NoBatch = "Please Enter Bordereaux Date and Batch No!", NoRange = "Please Enter Bordereaux Range!", NoVersion = "Please Enter Listing Version!",
        NoPayor = "Please Enter Payor!", NoGroup = "Please Enter Group Company!", NoRecord = "No record found!",
        MultiInvoice = "Different Invoice Found in Same Batch or Range, Please Search by Invoice No ! ", NoReportType = "Please select Report Type ! ";

    /// <inheritdoc />
    public async Task<(List<(string Code, string Name)> Payors, List<(string Code, string Name)> Groups)> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s);
    }

    /// <inheritdoc />
    public async Task<McoListResult> SearchAsync(McoListRequest r)
    {
        var rows = await Rows(r, null, null);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var princ = rows.Count(x => T(x, "MBMCRELCode") == "P");
        var invoices = rows.Select(x => T(x, "MBMCInvoiceNo")).Where(x => x.Length > 0).Distinct().Count();
        return new McoListResult(rows, princ, rows.Count - princ, !rows.Any(x => T(x, "MBMCInvoiceNo").Length > 0), invoices > 1);
    }

    /// <inheritdoc />
    public async Task<FileOutput> PrintAsync(McoListPrintRequest p, string user)
    {
        var r = p.Search;
        var search = await SearchAsync(r);
        if (search.MultipleInvoices) throw new BusinessRuleException(MultiInvoice);
        if (p.ReportType is not (0 or 1)) throw new BusinessRuleException(NoReportType);
        var byPayor = r.Version == 0;
        var selectedPayor = (p.SelectedPayor ?? "").Trim() is { Length: > 0 } sp ? sp : T(search.Rows[0], "PAYCode");
        var rows = await Rows(r, byPayor ? selectedPayor : null, p.SelectedBatch);
        var dept = r.Version == 2;
        if (p.ReportType == 1)
        {
            var mode = r.Criteria ?? 0;
            var type = mode switch
            {
                0 => $"BORDEREAUX DATE {new string(' ', 9)}:  {r.BordxDate}       Batch No :    {r.BatchNo}",
                1 => $"BORDEREAUX DATE {new string(' ', 9)}:  FROM    {r.From}      TO   {r.To}",
                _ => $"INVOICE NO. {new string(' ', 26)}:  {r.InvoiceNo}",
            };
            var payGrp = byPayor ? $"PAYOR {new string(' ', 38)}:  {r.Payor}" : $"GROUP COMPANY {new string(' ', 13)}:  {T(rows.FirstOrDefault(), "MBMGroupCompany")}";
            var sorted = rows.OrderBy(x => T(x, "MBMNumber"), StringComparer.Ordinal).ThenBy(x => T(x, "MBMCCoverID"), StringComparer.Ordinal).ToList();
            return exporter.Export(new ReportDocument(dept ? "MCO FEE LISTING BY DEPARTMENT" : "MCO FEE LISTING", ListColumns(dept), sorted.Cast<IReadOnlyDictionary<string, object?>>().ToList())
            {
                Subtitles = [payGrp.Trim(), type.Trim()], FileName = "MCOFeeListing",
            }, ReportFileType.Pdf);
        }

        // PrintMCOInvoice: the payor (by payor) or the group code of the first row; the company block always comes from GRPCompany (desktop bug).
        var payCode = byPayor ? selectedPayor : T(rows.FirstOrDefault(), "GRPCoCode");
        var withInvoice = rows.FirstOrDefault(x => T(x, "MBMCInvoiceNo").Length > 0);
        var invoiceNo = T(withInvoice, "MBMCInvoiceNo");
        var invoiceDate = withInvoice?.GetValueOrDefault("MBMCInvoiceDate") is DateTime d ? d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : "";
        await using var s = await sessions.OpenAsync();
        var company = await repository.CompanyAsync(s, payCode);
        if (p.Print && withInvoice is null)
        {
            await using var tx = await sessions.BeginAsync();
            invoiceNo = await repository.NextInvoiceNoAsync(tx);
            invoiceDate = clock.GetLocalNow().ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            foreach (var x in rows.Where(x => T(x, "MBMCInvoiceNo").Length == 0))
                await repository.StampAsync(tx, T(x, "MBMNumber"), T(x, "MBMCCoverID"), invoiceNo, user, x.GetValueOrDefault("MBMCMCOFee") is { } f ? Convert.ToDecimal(f, CultureInfo.InvariantCulture) : 0);
            await tx.CommitAsync();
        }
        var sortedInv = rows.OrderBy(x => x.GetValueOrDefault("MBMBordxDate") as DateTime?).ThenBy(x => T(x, "INSCode"), StringComparer.Ordinal).ToList();
        return exporter.Export(new ReportDocument("MCO INVOICE", InvoiceColumns(dept), sortedInv.Cast<IReadOnlyDictionary<string, object?>>().ToList())
        {
            Subtitles =
            [
                $"{payCode}  {T(company, "CompanyName")}", string.Join(", ", new[] { T(company, "Add1"), T(company, "Add2"), T(company, "Add3") }.Where(x => x.Length > 0)),
                $"Attn: {T(company, "Contact")}   Tel: {T(company, "TelNo")}   Fax: {T(company, "FaxNo")}", $"Invoice No: {invoiceNo}   Invoice Date: {invoiceDate}",
            ],
            FileName = invoiceNo.Length > 0 ? $"MCOInvoice_{invoiceNo}" : "MCOInvoice",
        }, ReportFileType.Pdf);
    }

    private static List<ReportColumn> ListColumns(bool dept) =>
    [
        new("MBMNumber", "Member No"), new("MBMCCoverID", "Cov"), new("MBMCRELCode", "Rel"), new("MBMName", "Name"), new("MBMCName", "Covered Person"),
        .. dept ? new[] { new ReportColumn("MBMDepartment", "Department") } : [], new("MBMBordxDate", "Bordx Date", ReportFormat.Date),
        new("MBMPolEffDate", "Eff Date", ReportFormat.Date), new("MBMPolExpDate", "Exp Date", ReportFormat.Date), new("MBMCMCOFee", "MCO Fee", ReportFormat.Money, true),
    ];

    private static List<ReportColumn> InvoiceColumns(bool dept) =>
    [
        new("MBMBordxDate", "Bordx Date", ReportFormat.Date), new("INSCode", "Ins"), new("MBMNumber", "Member No"), new("MBMCCoverID", "Cov"), new("MBMName", "Name"),
        .. dept ? new[] { new ReportColumn("MBMDepartment", "Department") } : [], new("MBMPolEffDate", "Eff Date", ReportFormat.Date),
        new("MBMPolExpDate", "Exp Date", ReportFormat.Date), new("MBMCMCOFee", "MCO Fee", ReportFormat.Money, true),
    ];

    /// <summary>cmdMCOSearch_Click's checks and MCOFeeListing (payor / batch overridden for the report).</summary>
    private async Task<List<Dictionary<string, object?>>> Rows(McoListRequest r, string? payor, string? batch)
    {
        if (r.Criteria is not (0 or 1 or 2)) throw new BusinessRuleException(NoCriteria);
        DateTime? bordx = null, from = null, to = null;
        int? batchNo = null;
        if (r.Criteria == 2)
        {
            if (string.IsNullOrWhiteSpace(r.InvoiceNo)) throw new BusinessRuleException(NoInvoice);
        }
        else if (r.Criteria == 0)
        {
            bordx = ClinicInvoiceRules.Date(r.BordxDate);
            if (bordx is null || string.IsNullOrWhiteSpace(r.BatchNo)) throw new BusinessRuleException(NoBatch);
            batchNo = int.TryParse((batch ?? r.BatchNo)!.Trim(), out var b) ? b : throw new BusinessRuleException(NoBatch);
        }
        else
        {
            from = ClinicInvoiceRules.Date(r.From);
            to = ClinicInvoiceRules.Date(r.To);
            if (from is null || to is null) throw new BusinessRuleException(NoRange);
        }
        if (r.Version is not (0 or 1 or 2)) throw new BusinessRuleException(NoVersion);
        string? pay = null, grp = null;
        if (r.Version == 0)
        {
            if (string.IsNullOrWhiteSpace(r.Payor)) throw new BusinessRuleException(NoPayor);
            pay = payor ?? Left(r.Payor, 2);
        }
        else
        {
            if (string.IsNullOrWhiteSpace(r.GroupCompany)) throw new BusinessRuleException(NoGroup);
            grp = Left(r.GroupCompany, 4);
        }
        await using var s = await sessions.OpenAsync();
        return await repository.SearchAsync(s, new McoListFilter(r.Criteria.Value, bordx, batchNo, from, to, r.InvoiceNo?.Trim(),
            (r.Invoiced ?? "N").Trim().StartsWith('Y') ? "Y" : "N", pay, grp));
    }

    /// <summary>Mid(x, 1, n): the code part of "code - name".</summary>
    private static string Left(string v, int n) => v.Trim().Length <= n ? v.Trim() : v.Trim()[..n].Trim();

    private static string T(IReadOnlyDictionary<string, object?>? r, string k) =>
        r is not null && r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";
}
