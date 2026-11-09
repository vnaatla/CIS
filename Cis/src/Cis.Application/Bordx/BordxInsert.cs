using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Bordx;

/// <summary>Insert Worksheet search (frmWSBordx).</summary>
public sealed record BordxWsRequest
{
#pragma warning disable CS1591
    /// <summary>0 INSURER-KIS, 1 ASO-KDC, 2 CREDIT NOTE-CN, 3 SIHAT JOHOR-KSJ, 4 VOID-KIV.</summary>
    public int? BordxType { get; init; }
    /// <summary>ASO-KDC options: 0 (optASO 0, pure ASO), 1 (optASO 1, BTBG), 2 (optASO 2, ASO type).</summary>
    public int? AsoOption { get; init; }
    /// <summary>"GRPCoName|GRPCoCode".</summary>
    public string? GroupCompany { get; init; }
    public bool Scan { get; init; }
    public string? CaseFrom { get; init; }
    public string? CaseTo { get; init; }
    public string? VisitFrom { get; init; }
    public string? VisitTo { get; init; }
    public bool NoVisitDate { get; init; }
    public string? Department { get; init; }
    /// <summary>0 All, 1 Specialised, 2 General Practisional, 3 Others.</summary>
    public int ClinicType { get; init; }
    public string? OpBatchFrom { get; init; }
    public string? OpBatchTo { get; init; }
    public string? Subsidiary { get; init; }
    /// <summary>C (Cashless) or R (Reimbursement).</summary>
    public string? CashType { get; init; }
    public string? Treatment { get; init; }
    public bool Omni { get; init; }
    public string? RegFrom { get; init; }
    public string? RegTo { get; init; }
    /// <summary>"code - name".</summary>
    public string? Clinic { get; init; }
    /// <summary>0 Health Connect, 1 Sihat Johor.</summary>
    public int Entity { get; init; }
    /// <summary>UOB list: 0 Outpatient, 1 flexCare (CboUOB; null = the normal list).</summary>
    public int? Uob { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class BordxWsRequestValidator : Common.RequestValidator<BordxWsRequest>
{
    /// <summary>Defines the rules.</summary>
    public BordxWsRequestValidator()
    {
        RuleFor(r => r.GroupCompany).MaximumLength(300);
        RuleFor(r => r.CaseFrom).MaximumLength(30);
        RuleFor(r => r.CaseTo).MaximumLength(30);
        RuleFor(r => r.Department).MaximumLength(200);
        RuleFor(r => r.Subsidiary).MaximumLength(200);
        RuleFor(r => r.Clinic).MaximumLength(300);
        RuleFor(r => r.OpBatchFrom).MaximumLength(30);
        RuleFor(r => r.OpBatchTo).MaximumLength(30);
        RuleFor(r => r.Treatment).MaximumLength(200);
    }
}

/// <summary>CmdBordx_Click: the search plus the ticked cases (in tick order) and the report type.</summary>
public sealed record BordxWsSaveRequest
{
    /// <summary>The search the cases were listed with.</summary>
    public BordxWsRequest Search { get; init; } = new();
    /// <summary>Ticked case numbers, in the order they were ticked.</summary>
    public List<string> Cases { get; init; } = [];
    /// <summary>0 By Group Company, 1 By Department.</summary>
    public int ReportType { get; init; }
}

/// <summary>Lengths.</summary>
public sealed class BordxWsSaveRequestValidator : Common.RequestValidator<BordxWsSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public BordxWsSaveRequestValidator() => RuleFor(r => r.Cases).Must(c => c.Count <= 20000 && c.All(x => x.Length <= 30));
}

/// <summary>The listed worksheets; each row carries the desktop's tick message (null = can be ticked).</summary>
public record BordxWsResult(List<Dictionary<string, object?>> Rows, int Total, decimal TotalAmount, bool Uob);

/// <summary>Data access.</summary>
public interface IBordxWsRepository
{
#pragma warning disable CS1591
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, BordxWsFilter filter);
    Task<(List<(string Code, string Name)> Groups, List<(string Code, string Name)> Treatments, List<(string Code, string Name)> Clinics)> LookupsAsync(IDbSession session);
    Task<(List<string> Departments, List<string> Subsidiaries)> CompanyAsync(IDbSession session, string company);
    Task<bool> IsBtbgAsync(IDbSession session, string plnCode);
    Task<string> NextNoAsync(IDbSession session, string payCode);
    Task<decimal> UpdateCaseAsync(IDbSession session, string casNumber, string bordxNo, int mode, string user);
    Task SaveBordxAsync(IDbSession session, string bordxNo, string payCode, string user, decimal amount, int claims, string company, string rptType, string subsidiary);
#pragma warning restore CS1591
}

/// <summary>cis.BordxWs_Search / SearchUob parameters.</summary>
public record BordxWsFilter(int BordxType, int IsPureAso, int? Uob, string? GroupCode, bool Scan, string? CaseFrom, string? CaseTo, DateTime? VisitFrom, DateTime? VisitTo,
    bool NoVisitDate, string? Department, int ClinicType, string? OpFrom, string? OpTo, string? Subsidiary, string? PayTo, string? Treatment, bool Omni,
    DateTime? RegFrom, DateTime? RegTo, string? Clinic, int Entity);

/// <summary>Claim Bordx > Worksheet > Insert Worksheet (frmWSBordx).</summary>
public interface IBordxWsService
{
    /// <summary>Form_Load drop-downs.</summary>
    Task<(List<(string Code, string Name)> Groups, List<(string Code, string Name)> Treatments, List<(string Code, string Name)> Clinics)> LookupsAsync();

    /// <summary>CboPayor_LostFocus: the company's departments and subsidiaries.</summary>
    Task<(List<string> Departments, List<string> Subsidiaries)> CompanyAsync(string? groupCompany);

    /// <summary>cmdSearch_Click / CboUOB_Click.</summary>
    Task<BordxWsResult> SearchAsync(BordxWsRequest request);

    /// <summary>CmdBordx_Click (after "Do you want to update the record?"): the new bordx no.</summary>
    Task<string> BordxAsync(BordxWsSaveRequest request, string user);
}

/// <inheritdoc />
public sealed class BordxWsService(IDbSessions sessions, IBordxWsRepository repository) : IBordxWsService
{
    /// <summary>Desktop texts.</summary>
    public const string NoGroup = "Please select a Group Company Name", NoWorksheet = "Please Enter Worksheet No", NoType = "Please select Bordx Type!",
        NoAso = "Please select Bordx Type whether ASO or BTBG..!", NoRecord = "No record found!", NoRecords = "Please select a record before proceeding!",
        NoneTicked = "Please select at least one record to be updated! ", NoUob = "Please select Flexcare or Outpatient Type!", Saved = "Record Saved, Your Bordx No is : ";

    /// <summary>The UOB companies whose lists are split into flexCare / outpatient.</summary>
    public static readonly string[] UobCompanies = ["UNITED OVERSEAS BANK (MALAYSIA) BHD", "UOB PROPERTIES BHD", "UNITED OVERSEAS BANK LIMITED"];

    /// <inheritdoc />
    public async Task<(List<(string Code, string Name)> Groups, List<(string Code, string Name)> Treatments, List<(string Code, string Name)> Clinics)> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s);
    }

    /// <inheritdoc />
    public async Task<(List<string> Departments, List<string> Subsidiaries)> CompanyAsync(string? groupCompany)
    {
        var name = (groupCompany ?? "").Split('|')[0].Trim();
        if (name.Length == 0) return ([], []);
        await using var s = await sessions.OpenAsync();
        return await repository.CompanyAsync(s, name);
    }

    /// <inheritdoc />
    public async Task<BordxWsResult> SearchAsync(BordxWsRequest r)
    {
        if (r.Scan)
        {
            if (string.IsNullOrWhiteSpace(r.GroupCompany)) throw new BusinessRuleException(NoGroup);
            if (string.IsNullOrWhiteSpace(r.CaseFrom)) throw new BusinessRuleException(NoWorksheet);
        }
        if (r.BordxType is not (>= 0 and <= 4)) throw new BusinessRuleException(NoType);
        if (r.BordxType == 1 && r.AsoOption is not (0 or 1 or 2)) throw new BusinessRuleException(NoAso);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, Filter(r));
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        var btbg = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        foreach (var plan in rows.Select(x => T(x, "PLNCode")).Distinct()) btbg[plan] = await repository.IsBtbgAsync(s, plan);
        foreach (var x in rows)
        {
            x["Name"] = T(x, "CoverID") == "00" ? T(x, "MBMName") : T(x, "MBMCName");
            x["Position"] = T(x, "MBMHISPolNo") switch
            {
                "03215666" => "AX - AXA AFFIN (NON-EXECUTIVE)", "03215661" => "AX - AXA AFFIN (EXECUTIVE AND ABOVE)", _ => T(x, "MBMPosition"),
            };
            x["TickError"] = TickError(x, btbg.GetValueOrDefault(T(x, "PLNCode")));
        }
        // The desktop's UOB test (operator precedence kept): one company and UOB Malaysia, or any first company that is UOB Properties / Limited.
        var first = T(rows[0], "MBMGroupCompany");
        var single = rows.All(x => T(x, "MBMGroupCompany") == first);
        var uob = (single && first == UobCompanies[0]) || first == UobCompanies[1] || first == UobCompanies[2];
        return new BordxWsResult(rows, rows.Count, rows.Sum(x => N(x, "CLMBordxAmt")), uob);
    }

    /// <summary>lstWorksheetTrackList_ItemCheck: the first failing rule's message, or null.</summary>
    public static string? TickError(IReadOnlyDictionary<string, object?> x, bool planIsBtbg)
    {
        if (x.GetValueOrDefault("DateVisit") is null) return "Date of Visit cannot be empty.";
        if (T(x, "CLNName").Length == 0) return "Clinic Name cannot be empty!";
        if (T(x, "MBMGroupCompany").Length == 0) return "Group Company cannot be empty!";
        if (T(x, "OPBatchNo").Length == 0) return "Batch Number is empty!";
        if (T(x, "CLNFFSStatus") == "Y" && Math.Abs(N(x, "CLMMCOProcFee")) == 0) return "FFS is empty!";
        if (T(x, "GRPMCOStatus").Length == 0) return "Group Company's FFS Status is empty!";
        if (N(x, "CLMBordxAmt") == 0) return "Please check Bordx Amount!";
        if (T(x, "INVPayTo") == "C" && Money(x, "CLMTotalActual") != Money(x, "INVBillAmt")) return "Difference in Bill Amount and Payable to Clinic Amount!";
        if (N(x, "CLMCoPayColl") == 0 && Money(x, "CLMBordxAmt") != Money(x, "CLMTotalActual") && !planIsBtbg) return "Please check Bill Amount & Bordx Amount!";
        return null;
    }

    /// <inheritdoc />
    public async Task<string> BordxAsync(BordxWsSaveRequest req, string user)
    {
        var r = req.Search;
        if (req.Cases.Count == 0) throw new BusinessRuleException(NoneTicked);
        if (r.BordxType is not (>= 0 and <= 4)) throw new BusinessRuleException(NoType);
        var search = await SearchAsync(r);
        var listed = search.Rows;
        if (search.Uob && r.Uob is null) throw new BusinessRuleException(NoUob);
        var byCase = listed.GroupBy(x => T(x, "CASNumber")).ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        var ticked = req.Cases.Select(c => c.Trim()).Where(byCase.ContainsKey).ToList();
        if (ticked.Count == 0) throw new BusinessRuleException(NoneTicked);
        var refused = ticked.Select(c => T(byCase[c], "TickError")).FirstOrDefault(e => e.Length > 0);
        if (refused is not null) throw new BusinessRuleException(refused);
        var payCode = r.BordxType switch
        {
            0 => "IS", 1 => "DC", 2 => "CN", 3 when r.Entity == 1 => "SJ", 4 => "IV", _ => throw new BusinessRuleException(NoType),
        };
        var company = (r.GroupCompany ?? "").Split('|')[0].Trim();
        await using var tx = await sessions.BeginAsync();
        // CheckIsBTBG uses the plan of the last ticked case (the desktop's PLNCode); INSURER-KIS is always treated as BTBG.
        var btbg = r.BordxType == 0 || await repository.IsBtbgAsync(tx, T(byCase[ticked[^1]], "PLNCode"));
        var mode = payCode == "IV" ? 1 : (payCode is "DC" or "CN") && btbg ? 2 : 0;
        var bordxNo = await repository.NextNoAsync(tx, payCode);
        decimal total = 0;
        foreach (var c in ticked) total += await repository.UpdateCaseAsync(tx, c, bordxNo, mode, user);
        await repository.SaveBordxAsync(tx, bordxNo, company.Length > 0 ? "DC" : "", user, total, ticked.Count, company, req.ReportType == 1 ? "1" : "",
            (r.Subsidiary ?? "").Trim());
        await tx.CommitAsync();
        return bordxNo;
    }

    private static BordxWsFilter Filter(BordxWsRequest r)
    {
        var clinic = (r.Clinic ?? "").Trim();
        if (clinic.IndexOf('-') is > 0 and var i) clinic = clinic[..i].Trim();
        return new BordxWsFilter(r.BordxType!.Value, r.AsoOption switch { 0 => 1, 2 => 2, _ => 0 }, r.Uob,
            (r.GroupCompany ?? "").Split('|') is { Length: > 1 } g && g[1].Trim().Length > 0 ? g[1].Trim() : null, r.Scan, Blank(r.CaseFrom), r.Scan ? null : Blank(r.CaseTo),
            ClinicInvoiceRules.Date(r.VisitFrom), ClinicInvoiceRules.Date(r.VisitTo), r.NoVisitDate, Blank(r.Department), r.ClinicType, Blank(r.OpBatchFrom),
            Blank(r.OpBatchTo), Blank(r.Subsidiary), (r.CashType ?? "").Trim() switch { "" => null, var c when c.StartsWith('C') => "C", _ => "M" },
            Blank(r.Treatment) is { } t ? t[..Math.Min(3, t.Length)] : null, r.Omni, ClinicInvoiceRules.Date(r.RegFrom), ClinicInvoiceRules.Date(r.RegTo),
            clinic.Length > 0 ? clinic : null, r.Entity);
    }

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull && decimal.TryParse(Convert.ToString(v, CultureInfo.InvariantCulture), NumberStyles.Any,
            CultureInfo.InvariantCulture, out var d) ? d : 0;

    /// <summary>The desktop compares the list's "#,##0.00" texts (blank when null).</summary>
    private static string Money(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? N(r, k).ToString("#,##0.00", CultureInfo.InvariantCulture) : "";
}
