using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Mco;

/// <summary>Refund Listing filters (frmMCORefundListing).</summary>
public sealed record McoRefundRequest
{
    /// <summary>0 by Payor, 1 by Group Company, 2 by Department.</summary>
    public int? Version { get; init; }
    /// <summary>0 Endt. Eff Date, 1 Endt. Bordx Date.</summary>
    public int? DateSelection { get; init; }
    /// <summary>dd/mm/yyyy (invalid dates are ignored, as the desktop did).</summary>
    public string? From { get; init; }
    /// <summary>dd/mm/yyyy.</summary>
    public string? To { get; init; }
    /// <summary>"code - name".</summary>
    public string? Payor { get; init; }
    /// <summary>Group company name.</summary>
    public string? GroupCompany { get; init; }
    /// <summary>Department.</summary>
    public string? Department { get; init; }
    /// <summary>N (no credit note yet) or Y.</summary>
    public string? Status { get; init; }
}

/// <summary>Lengths.</summary>
public sealed class McoRefundRequestValidator : Common.RequestValidator<McoRefundRequest>
{
    /// <summary>Defines the rules.</summary>
    public McoRefundRequestValidator()
    {
        RuleFor(r => r.Payor).MaximumLength(200);
        RuleFor(r => r.GroupCompany).MaximumLength(200);
        RuleFor(r => r.Department).MaximumLength(200);
    }
}

/// <summary>A listed row (principal or covered person).</summary>
public sealed record McoRefundRow
{
#pragma warning disable CS1591
    public string Member { get; init; } = "";
    public string Cover { get; init; } = "";
    public bool Principal { get; init; }
    public string Name { get; init; } = "";
    public string Ic { get; init; } = "";
    public DateTime? BordxDate { get; init; }
    public DateTime? EffDate { get; init; }
    public DateTime? ExpDate { get; init; }
    public string DaysInsured { get; init; } = "";
    public DateTime? EndtEffDate { get; init; }
    public DateTime? EndtBordxDate { get; init; }
    public string DaysUtilised { get; init; } = "";
    public decimal? McoFee { get; init; }
    public decimal? McoUtilised { get; init; }
    public decimal? McoRefund { get; init; }
    public decimal McoUtilisedFee { get; init; }
    public string InvoiceNo { get; init; } = "";
    public DateTime? InvoiceDate { get; init; }
    public string Cases { get; init; } = "";
    public string CreditNote { get; init; } = "";
    public DateTime? CreditNoteDate { get; init; }
#pragma warning restore CS1591
}

/// <summary>The listing and its totals (principals, covered persons, refunds, utilised fees).</summary>
public record McoRefundResult(List<McoRefundRow> Rows, int Principals, int Supplementaries, decimal PrincipalRefund, decimal UtilisedFee, decimal SupplementaryRefund,
    decimal TotalRefund);

/// <summary>Data access.</summary>
public interface IMcoRefundRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> PrincipalsAsync(IDbSession session, McoRefundFilter filter, string noted);
    Task<List<Dictionary<string, object?>>> SupplementariesAsync(IDbSession session, McoRefundFilter filter, IReadOnlyList<string> members);
    Task<(List<(string Code, string Name)> Payors, List<string> Companies, List<string> Departments)> LookupsAsync(IDbSession session, string? company);
    Task<Dictionary<string, object?>?> CompanyAsync(IDbSession session, string? payor, string? company);
    Task<string> NextCreditNoteAsync(IDbSession session);
    Task StampAsync(IDbSession session, string member, string creditNote, string user);
#pragma warning restore CS1591
}

/// <summary>cis.McoRefund_Search parameters.</summary>
public record McoRefundFilter(string? Payor, string? Company, string? Department, bool ByBordx, DateTime? From, DateTime? To);

/// <summary>MCO Fees > Refund Listing (frmMCORefundListing).</summary>
public interface IMcoRefundService
{
    /// <summary>Drop-downs (departments of a company).</summary>
    Task<(List<(string Code, string Name)> Payors, List<string> Companies, List<string> Departments)> LookupsAsync(string? company);

    /// <summary>cmdSearch_Click.</summary>
    Task<McoRefundResult> SearchAsync(McoRefundRequest request);

    /// <summary>CmdPrint_Click: the listing as Excel.</summary>
    Task<FileOutput> ExportAsync(McoRefundRequest request);

    /// <summary>cmdPrintCRNote_Click: the credit note (status N numbers it and stamps the members).</summary>
    Task<FileOutput> CreditNoteAsync(McoRefundRequest request, string user);
}

/// <inheritdoc />
public sealed class McoRefundService(IDbSessions sessions, IMcoRefundRepository repository, IReportExporter exporter, TimeProvider clock) : IMcoRefundService
{
    /// <summary>Desktop texts.</summary>
    public const string NoVersion = "Please Enter Listing Version!", NoEffDate = "Please Enter Endt. Eff Date!", NoBordxDate = "Please Enter Endt. Bordx Date!",
        NoPayor = "Please Select a Payor!", NoCompany = "Please Select a Group Company!", NoDepartment = "Please Select a Department!", NoRecord = "No record found!",
        NothingToPrint = "Report cannot be printed without any record.", NoRecords = "Please Select Records!";

    /// <inheritdoc />
    public async Task<(List<(string Code, string Name)> Payors, List<string> Companies, List<string> Departments)> LookupsAsync(string? company)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s, string.IsNullOrWhiteSpace(company) ? null : company.Trim());
    }

    /// <inheritdoc />
    public async Task<McoRefundResult> SearchAsync(McoRefundRequest r)
    {
        Check(r);
        var result = await Listing(r);
        return result.Rows.Count == 0 ? throw new NotFoundException(NoRecord) : result;
    }

    /// <inheritdoc />
    public async Task<FileOutput> ExportAsync(McoRefundRequest r)
    {
        Check(r);
        var result = await Listing(r);
        if (result.Rows.Count == 0) throw new BusinessRuleException(NothingToPrint);
        return exporter.Export(new ReportDocument("MCO REFUND LISTING", Columns, result.Rows.Select(Dict).ToList()) { FileName = "MCORefundListing" }, ReportFileType.Excel);
    }

    /// <inheritdoc />
    public async Task<FileOutput> CreditNoteAsync(McoRefundRequest r, string user)
    {
        if (r.Version is not (0 or 1 or 2)) throw new BusinessRuleException(NoVersion);
        var result = await Listing(r);
        if (result.Rows.Count == 0) throw new BusinessRuleException(NoRecords);
        if (r.DateSelection is not (0 or 1)) throw new BusinessRuleException("Please Select Date Selection!");
        if (r.Version == 0 && string.IsNullOrWhiteSpace(r.Payor)) throw new BusinessRuleException("Please Select Payor!");
        if (r.Version is 1 or 2 && string.IsNullOrWhiteSpace(r.GroupCompany)) throw new BusinessRuleException("Please Select Group Company!");
        var newNote = !(r.Status ?? "N").Trim().StartsWith('Y');
        await using var s = await sessions.OpenAsync();
        var company = await repository.CompanyAsync(s, r.Version == 0 ? Code(r.Payor) : null, r.Version == 0 ? null : r.GroupCompany?.Trim());
        string note;
        DateTime? noteDate;
        if (newNote)
        {
            await using var tx = await sessions.BeginAsync();
            note = await repository.NextCreditNoteAsync(tx);
            foreach (var m in result.Rows.Select(x => x.Member).Distinct()) await repository.StampAsync(tx, m, note, user);
            await tx.CommitAsync();
            noteDate = clock.GetLocalNow().DateTime.Date;
        }
        else
        {
            var first = result.Rows.FirstOrDefault(x => x.Principal);
            note = first?.CreditNote ?? "";
            noteDate = first?.CreditNoteDate;
        }
        var total = result.Rows.Sum(x => x.McoRefund ?? 0);
        string S(string k) => Convert.ToString(company?.GetValueOrDefault(k), CultureInfo.InvariantCulture)?.Trim() ?? "";
        var heading = r.Version == 2 ? $"Dept.  {r.Department?.Trim()}" : S("Name");
        return exporter.Export(new ReportDocument("MCO CREDIT NOTE", Columns, result.Rows.Select(Dict).ToList())
        {
            Subtitles =
            [
                heading, S("Add1"), S("Add2"), S("Add3"), $"Tel: {S("Tel")}   Contact: {S("Contact")}",
                $"Credit Note No: {note}   Date: {noteDate?.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}",
                $"Total Refund: {total.ToString("#,##0.00", CultureInfo.InvariantCulture)}",
            ],
            FileName = $"MCOCreditNote_{note}",
        }, ReportFileType.Excel);
    }

    private static readonly IReadOnlyList<ReportColumn> Columns =
    [
        new("Member", "Member No"), new("Cover", "Cov"), new("Name", "Name"), new("Ic", "IC No"), new("BordxDate", "Bordx Date", ReportFormat.Date),
        new("EffDate", "Eff Date", ReportFormat.Date), new("ExpDate", "Exp Date", ReportFormat.Date), new("DaysInsured", "Days Insured"),
        new("EndtEffDate", "Endt Eff Date", ReportFormat.Date), new("EndtBordxDate", "Endt Bordx Date", ReportFormat.Date), new("DaysUtilised", "Days Utilised"),
        new("McoFee", "MCO Fee", ReportFormat.Money), new("McoUtilised", "MCO Utilised", ReportFormat.Money), new("McoRefund", "MCO Refund", ReportFormat.Money, true),
        new("McoUtilisedFee", "MCO Fee Utilised", ReportFormat.Money, true), new("InvoiceNo", "Invoice No"), new("InvoiceDate", "Invoice Date", ReportFormat.Date),
        new("Cases", "Cases"), new("CreditNote", "CR Note"), new("CreditNoteDate", "CR Note Date", ReportFormat.Date),
    ];

    private static IReadOnlyDictionary<string, object?> Dict(McoRefundRow x) => new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
    {
        ["Member"] = x.Member, ["Cover"] = x.Cover, ["Name"] = x.Name, ["Ic"] = x.Ic, ["BordxDate"] = x.BordxDate, ["EffDate"] = x.EffDate, ["ExpDate"] = x.ExpDate,
        ["DaysInsured"] = x.DaysInsured, ["EndtEffDate"] = x.EndtEffDate, ["EndtBordxDate"] = x.EndtBordxDate, ["DaysUtilised"] = x.DaysUtilised, ["McoFee"] = x.McoFee,
        ["McoUtilised"] = x.McoUtilised, ["McoRefund"] = x.McoRefund, ["McoUtilisedFee"] = x.McoUtilisedFee, ["InvoiceNo"] = x.InvoiceNo, ["InvoiceDate"] = x.InvoiceDate,
        ["Cases"] = x.Cases, ["CreditNote"] = x.CreditNote, ["CreditNoteDate"] = x.CreditNoteDate,
    };

    /// <summary>cmdSearch_Click's checks (the desktop tests the To date twice and never the From date).</summary>
    private static void Check(McoRefundRequest r)
    {
        if (r.Version is not (0 or 1 or 2)) throw new BusinessRuleException(NoVersion);
        if (r.DateSelection is not (0 or 1)) return;
        if (string.IsNullOrWhiteSpace(r.To)) throw new BusinessRuleException(r.DateSelection == 0 ? NoEffDate : NoBordxDate);
        if (r.Version == 0 && string.IsNullOrWhiteSpace(r.Payor)) throw new BusinessRuleException(NoPayor);
        if (r.Version is 1 or 2 && string.IsNullOrWhiteSpace(r.GroupCompany)) throw new BusinessRuleException(NoCompany);
        if (r.Version == 2 && string.IsNullOrWhiteSpace(r.Department)) throw new BusinessRuleException(NoDepartment);
    }

    /// <summary>SearchRecord + MCORefundListing + MCORefundListingSupp.</summary>
    private async Task<McoRefundResult> Listing(McoRefundRequest r)
    {
        var filter = new McoRefundFilter(Code(r.Payor), Blank(r.GroupCompany), Blank(r.Department), r.DateSelection == 1,
            ClinicInvoiceRules.Date(r.From), ClinicInvoiceRules.Date(r.To));
        await using var s = await sessions.OpenAsync();
        var principals = await repository.PrincipalsAsync(s, filter, (r.Status ?? "N").Trim().StartsWith('Y') ? "Y" : "N");
        var family = principals.Where(p => T(p, "INSType") is "F" or "H").Select(p => T(p, "MBMNumber")).Distinct().ToList();
        var supps = family.Count == 0 ? [] : await repository.SupplementariesAsync(s, filter, family);
        var rows = new List<McoRefundRow>();
        decimal principalRefund = 0, utilisedFee = 0, suppRefund = 0;
        int suppCount = 0;
        foreach (var p in principals)
        {
            var row = Row(p, true, T(p, "INSType") is "F" or "H" ? 50 : 25);
            rows.Add(row);
            principalRefund += row.McoRefund ?? 0;
            utilisedFee += row.McoUtilisedFee;
            if (T(p, "INSType") is not ("F" or "H")) continue;
            foreach (var c in supps.Where(x => T(x, "MBMNumber") == row.Member))
            {
                var sr = Row(c, false, 50);
                rows.Add(sr);
                suppCount++;
                suppRefund += sr.McoRefund ?? 0;
            }
        }
        return new McoRefundResult(rows, principals.Count, suppCount, principalRefund, utilisedFee, suppRefund, principalRefund + suppRefund);
    }

    /// <summary>One list item (the name is shown only for cancelled / pending members, as the desktop fills it only then).</summary>
    private static McoRefundRow Row(Dictionary<string, object?> x, bool principal, decimal threshold)
    {
        var name = T(x, principal ? "MBMName" : "MBMCName");
        var status = T(x, "MBMCLNStatus");
        var fee = N(x, principal ? "MBMCLNMCOFee" : "MBMCCLNMCO");
        var utilised = N(x, "MBMMCOUtilised");
        var cases = x.GetValueOrDefault("TotalCase") is { } tc && Convert.ToInt32(tc, CultureInfo.InvariantCulture) > 0 ? $"This member had {tc} Case(s)" : "";
        return new McoRefundRow
        {
            Member = T(x, "MBMNumber"), Cover = T(x, principal ? "MBMCOVID" : "MBMCCoverID"), Principal = principal,
            Name = status == "C" ? $"{name} (Cancelled)" : status == "P" ? $"{name} (Pending)" : "", Ic = T(x, principal ? "MBMIcBcPp" : "MBMCICBCPPNo"),
            BordxDate = D(x, principal ? "MBMBordxDate" : "MBMCBordxDate"), EffDate = D(x, "MBMPolEffDate"), ExpDate = D(x, "MBMPolExpDate"),
            DaysInsured = T(x, "DayInsured"), EndtEffDate = D(x, "MBMCLNEndtEffDate"), EndtBordxDate = D(x, "MBMCLNEndtBordxDate"), DaysUtilised = T(x, "DayUtilised"),
            McoFee = fee, McoUtilised = utilised, McoRefund = N(x, "MBMMCORefund"),
            McoUtilisedFee = utilised is null ? 0 : utilised <= threshold ? (fee ?? 0) - threshold : utilised.Value,
            InvoiceNo = T(x, "MBMCInvoiceNo"), InvoiceDate = D(x, "MBMCInvoiceDate"), Cases = cases,
            CreditNote = T(x, principal ? "MBMCLNCRNote" : "MBMCCLNCRNote"), CreditNoteDate = D(x, principal ? "MBMCLNCRNoteDate" : "MBMCCLNCRNoteDate"),
        };
    }

    private static string? Code(string? payor) => string.IsNullOrWhiteSpace(payor) ? null : payor.Trim().Length <= 2 ? payor.Trim() : payor.Trim()[..2];

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal? N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToDecimal(v, CultureInfo.InvariantCulture) : null;

    private static DateTime? D(IReadOnlyDictionary<string, object?> r, string k) => r.GetValueOrDefault(k) as DateTime?;
}
