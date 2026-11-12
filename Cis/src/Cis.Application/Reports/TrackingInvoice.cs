using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Invoice;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Reports;

/// <summary>Tracking > Invoice search (FrmInvoiceTracking SearchRecord). Text fields as typed; combos "code - text"; the *Blank flags are the "empty" check boxes.</summary>
public sealed record InvTrkRequest
{
#pragma warning disable CS1591
    public string? CaseFrom { get; init; }
    public string? CaseTo { get; init; }
    public bool CaseBlank { get; init; }
    public string? PayTo { get; init; }
    public bool PayToBlank { get; init; }
    public string? Clinic { get; init; }
    public bool PayeeBlank { get; init; }
    public string? Payor { get; init; }
    public bool PayorBlank { get; init; }
    public string? GroupCompany { get; init; }
    public bool GroupCompanyBlank { get; init; }
    public string? Patient { get; init; }
    public bool PatientBlank { get; init; }
    public string? MemberFrom { get; init; }
    public string? MemberTo { get; init; }
    public bool MemberBlank { get; init; }
    public string? InvoiceFrom { get; init; }
    public string? InvoiceTo { get; init; }
    public bool InvoiceBlank { get; init; }
    public string? BordxFrom { get; init; }
    public string? BordxTo { get; init; }
    public bool BordxBlank { get; init; }
    /// <summary>Case statuses ticked: O, D, C.</summary>
    public List<string> Statuses { get; init; } = [];
    public bool Pending { get; init; }
    /// <summary>Tracking (Jupiter): the single case status combo (O / C / D).</summary>
    public string? Status { get; init; }
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
    public string? AccMonthFrom { get; init; }
    public string? AccMonthTo { get; init; }
    public bool AccMonthBlank { get; init; }
    public string? AccYearFrom { get; init; }
    public string? AccYearTo { get; init; }
    public bool AccYearBlank { get; init; }
    /// <summary>"0 - Specialised", "1 - General Practisional", "2 - Others".</summary>
    public string? ClinicType { get; init; }
    public string? OpBatchFrom { get; init; }
    public string? OpBatchTo { get; init; }
    public string? OpBatchDateFrom { get; init; }
    public string? OpBatchDateTo { get; init; }
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
    /// <summary>0 Health Connect, 1 Sihat Johor.</summary>
    public int Entity { get; init; }
    /// <summary>null, 0 OutPatient, 1 flexCare.</summary>
    public int? Flex { get; init; }
    /// <summary>"", "Yes", "No" (with Pay To).</summary>
    public string? Paid { get; init; }
#pragma warning restore CS1591
}

/// <summary>Lengths.</summary>
public sealed class InvTrkRequestValidator : Common.RequestValidator<InvTrkRequest>
{
    /// <summary>Defines the rules.</summary>
    public InvTrkRequestValidator()
    {
        RuleFor(r => r.GroupCompany).MaximumLength(200);
        RuleFor(r => r.Patient).MaximumLength(200);
        RuleFor(r => r.Clinic).MaximumLength(300);
        RuleFor(r => r.Payor).MaximumLength(300);
        RuleFor(r => r.Subsidiary).MaximumLength(200);
        RuleFor(r => r.Department).MaximumLength(200);
        RuleFor(r => r.Statuses).Must(s => s.Count <= 4 && s.All(x => x is "O" or "D" or "C"));
        RuleFor(r => r.CaseFrom).MaximumLength(20);
        RuleFor(r => r.CaseTo).MaximumLength(20);
        RuleFor(r => r.MemberFrom).MaximumLength(30);
        RuleFor(r => r.MemberTo).MaximumLength(30);
        RuleFor(r => r.InvoiceFrom).MaximumLength(50);
        RuleFor(r => r.InvoiceTo).MaximumLength(50);
        RuleFor(r => r.BordxFrom).MaximumLength(20);
        RuleFor(r => r.BordxTo).MaximumLength(20);
        RuleFor(r => r.PlanFrom).MaximumLength(20);
        RuleFor(r => r.PlanTo).MaximumLength(20);
        RuleFor(r => r.IcFrom).MaximumLength(30);
        RuleFor(r => r.IcTo).MaximumLength(30);
        RuleFor(r => r.OpBatchFrom).MaximumLength(30);
        RuleFor(r => r.OpBatchTo).MaximumLength(30);
        RuleFor(r => r.OmniRefFrom).MaximumLength(30);
        RuleFor(r => r.OmniRefTo).MaximumLength(30);
    }
}

/// <summary>The list and its totals (invoice amount, approved, BTBG).</summary>
public sealed record InvTrkResult(List<Dictionary<string, object?>> Rows, decimal Total, decimal Approved, decimal Btbg);

/// <summary>Data access.</summary>
public interface IInvTrkRepository
{
#pragma warning disable CS1591
    Task<List<string>> LookupAsync(IDbSession session, string kind, string? key);
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, IReadOnlyDictionary<string, object?> parameters, bool jupiter);
#pragma warning restore CS1591
}

/// <summary>Tracking > Invoice (FrmInvoiceTracking) and Tracking (Jupiter) (FrmInvoiceTracking2, same search on the Jupiter server; here the same databases).</summary>
public interface IInvTrkService
{
    /// <summary>Combos: clinics, payors, companies (of a payor), subsidiaries / departments (of a group company).</summary>
    Task<List<string>> LookupAsync(string kind, string? key);

    /// <summary>cmdSearch_Click; <paramref name="jupiter"/> = FrmInvoiceTracking2 (ClaimBordxModule case 1).</summary>
    Task<InvTrkResult> SearchAsync(InvTrkRequest request, bool jupiter = false);
}

/// <inheritdoc />
public sealed class InvTrkService(IDbSessions sessions, IInvTrkRepository repository) : IInvTrkService
{
    /// <summary>Desktop texts.</summary>
    public const string NoRecord = "No record found!", BadDate = "Invalid Date Format! ";

    private static readonly string[] Kinds = ["clinics", "payors", "companies", "subsidiaries", "departments"];

    /// <inheritdoc />
    public async Task<List<string>> LookupAsync(string kind, string? key)
    {
        if (!Kinds.Contains(kind)) throw new NotFoundException("Unknown list.");
        await using var s = await sessions.OpenAsync();
        return await repository.LookupAsync(s, kind, string.IsNullOrWhiteSpace(key) ? null : key.Trim());
    }

    /// <inheritdoc />
    public async Task<InvTrkResult> SearchAsync(InvTrkRequest r, bool jupiter = false)
    {
        var p = Parameters(r);
        if (jupiter) p["status"] = string.IsNullOrWhiteSpace(r.Status) ? null : r.Status.Trim()[..1];
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, p, jupiter);
        if (rows.Count == 0) throw new NotFoundException(NoRecord);
        decimal total = 0, approved = 0, btbg = 0;
        foreach (var x in rows)
        {
            Shape(x);
            total += N(x, "INVAmount");
            approved += N(x, "CLMFinalApprove");
            btbg += N(x, "CLMBTBGAmt");
        }
        return new InvTrkResult(rows, total, approved, btbg);
    }

    /// <summary>InvoiceTrackingListing's derived columns: days (0 shown as 1), the clinic (clinic name when paid to clinic, else the non-panel clinic), reason 5 (the
    /// diagnosis for reasons 199 / 299), and the fee / medicine / injection sums.</summary>
    public static void Shape(Dictionary<string, object?> x)
    {
        var days = T(x, "NoofDays");
        x["DaysText"] = days == "0" ? "1" : days;
        x["ClinicText"] = T(x, "INVPayTo") == "C" ? T(x, "CLNName") : T(x, "NPClinic");
        var r5 = T(x, "CASTreatReason5");
        x["Reason5Text"] = r5.Length == 0 ? "" : r5 is not ("199" or "299") ? T(x, "ReasonDesc5") : T(x, "CASDiagnosis5");
        x["FeeTotal"] = N(x, "GPFee") + N(x, "SPFee");
        x["MedTotal"] = N(x, "GPLTMed") + N(x, "SPLTMed") + N(x, "GPMed") + N(x, "SPMed");
        x["InjTotal"] = N(x, "GPInj") + N(x, "SPInj");
    }

    private static Dictionary<string, object?> Parameters(InvTrkRequest r)
    {
        static string? B(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        static string? Code(string? v)
        {
            var t = B(v);
            return t is null ? null : (t.Contains('-') ? t[..t.IndexOf('-')] : t[..Math.Min(4, t.Length)]).Trim();
        }
        static int? I(string? v) => int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out var i) ? i : null;
        static decimal? M(string? v) => B(v) is { } t ? decimal.TryParse(t, NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d
            : throw new BusinessRuleException("Invalid amount.") : null;
        var payTo = B(r.PayTo)?[..1];
        var statuses = r.Statuses.Count == 0 ? null : string.Concat(r.Statuses);
        return new Dictionary<string, object?>
        {
            ["caseFrom"] = B(r.CaseFrom), ["caseTo"] = B(r.CaseTo), ["caseBlank"] = r.CaseBlank, ["payTo"] = payTo, ["payToBlank"] = r.PayToBlank,
            ["clinic"] = Code(r.Clinic), ["payeeBlank"] = r.PayeeBlank, ["payor"] = Code(r.Payor), ["payorBlank"] = r.PayorBlank,
            ["company"] = B(r.GroupCompany), ["companyBlank"] = r.GroupCompanyBlank, ["patient"] = B(r.Patient), ["patientBlank"] = r.PatientBlank,
            ["mbmFrom"] = B(r.MemberFrom), ["mbmTo"] = B(r.MemberTo), ["mbmBlank"] = r.MemberBlank, ["billFrom"] = B(r.InvoiceFrom), ["billTo"] = B(r.InvoiceTo),
            ["billBlank"] = r.InvoiceBlank, ["bordxFrom"] = B(r.BordxFrom), ["bordxTo"] = B(r.BordxTo), ["bordxBlank"] = r.BordxBlank, ["statuses"] = statuses,
            ["pending"] = r.Pending, ["claimability"] = B(r.Claimability)?[..1], ["planFrom"] = B(r.PlanFrom), ["planTo"] = B(r.PlanTo), ["planBlank"] = r.PlanBlank,
            ["bordxDateFrom"] = D(r.BordxDateFrom), ["bordxDateTo"] = D(r.BordxDateTo), ["icFrom"] = B(r.IcFrom), ["icTo"] = B(r.IcTo), ["icBlank"] = r.IcBlank,
            ["visitFrom"] = D(r.VisitFrom), ["visitTo"] = D(r.VisitTo), ["visitBlank"] = r.VisitBlank, ["regFrom"] = D(r.RegisterFrom), ["regTo"] = D(r.RegisterTo),
            ["regBlank"] = r.RegisterBlank, ["billDateFrom"] = D(r.InvoiceDateFrom), ["billDateTo"] = D(r.InvoiceDateTo), ["billDateBlank"] = r.InvoiceDateBlank,
            ["enteredFrom"] = D(r.EnteredFrom), ["enteredTo"] = D(r.EnteredTo), ["enteredBlank"] = r.EnteredBlank, ["amountFrom"] = M(r.AmountFrom),
            ["amountTo"] = M(r.AmountTo), ["amountBlank"] = r.AmountBlank, ["accMonthFrom"] = I(r.AccMonthFrom), ["accMonthTo"] = I(r.AccMonthTo),
            ["accMonthBlank"] = r.AccMonthBlank, ["accYearFrom"] = I(r.AccYearFrom), ["accYearTo"] = I(r.AccYearTo), ["accYearBlank"] = r.AccYearBlank,
            ["clinicType"] = B(r.ClinicType)?[..1], ["opBatchFrom"] = B(r.OpBatchFrom), ["opBatchTo"] = B(r.OpBatchTo), ["opBatchDateFrom"] = D(r.OpBatchDateFrom),
            ["opBatchDateTo"] = D(r.OpBatchDateTo), ["opBatchBlank"] = r.OpBatchBlank, ["ffs"] = B(r.Ffs)?[..1], ["omni"] = r.Omni, ["tims"] = r.Tims,
            ["subsidiary"] = B(r.Subsidiary), ["department"] = B(r.Department), ["omniRefFrom"] = B(r.OmniRefFrom), ["omniRefTo"] = B(r.OmniRefTo),
            ["omniRefBlank"] = r.OmniRefBlank, ["omniSent"] = B(r.OmniSent)?[..1], ["sj"] = r.Entity == 1, ["flex"] = r.Flex, ["paid"] = B(r.Paid)?[..1],
        };
    }

    private static DateTime? D(string? v) => string.IsNullOrWhiteSpace(v) ? null : ClinicInvoiceRules.Date(v) ?? throw new BusinessRuleException(BadDate);

    private static string T(IReadOnlyDictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null && v is not DBNull ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    private static decimal N(IReadOnlyDictionary<string, object?> r, string k) =>
        decimal.TryParse(T(r, k), NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0;
}
