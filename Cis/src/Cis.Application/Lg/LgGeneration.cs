using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Files;
using Cis.Application.Invoice;
using Cis.Application.Membership;
using Cis.Domain.Errors;

namespace Cis.Application.Lg;

/// <summary>A member / cover found by CLNLG1.</summary>
public sealed record LgMemberRow
{
#pragma warning disable CS1591
    public string Member { get; init; } = "";
    public string Cover { get; init; } = "";
    public string Name { get; init; } = "";
    public string Ic { get; init; } = "";
    public string GroupCompany { get; init; } = "";
    public string StaffCategory { get; init; } = "";
    public string ClinicStatus { get; init; } = "";
    public string Plan { get; init; } = "";
    public string PlanLgRemarks { get; init; } = "";
    public DateTime? PolicyExpiry { get; init; }
    public string PlanSpecialRemark { get; init; } = "";
    public bool Aso { get; init; }
    public string ClinicalRemarks { get; init; } = "";
    public string Category { get; init; } = "";
    public string Subsidiary { get; init; } = "";
#pragma warning restore CS1591
}

/// <summary>SearchRecords: the covers and, for the listed companies, the warranty message.</summary>
public record LgSearchResult(List<LgMemberRow> Rows, string? Warranty);

/// <summary>lstMBM_Click: limits, ValidateLG and the member's LGs.</summary>
/// <param name="Member">The cover clicked.</param>
/// <param name="Limits">CheckMBMLimit labels.</param>
/// <param name="Status">Member status (C when the principal's clinic status is C).</param>
/// <param name="ValidLg">IsValidLg.</param>
/// <param name="GpSpShared">SP_ValidateLG1 PLNGPSPShared.</param>
/// <param name="AfterLimitExInd">SP_ValidateLG1 LGIssuedAfterLimitExInd.</param>
/// <param name="Messages">Messages the desktop showed (M230 balance; "Member is not active.").</param>
/// <param name="Lgs">LGlst2020 rows.</param>
public record LgMemberInfo(LgMemberRow Member, LgLimits Limits, string Status, bool ValidLg, string GpSpShared, string AfterLimitExInd, List<string> Messages,
    List<Dictionary<string, object?>> Lgs);

/// <summary>lstLG_Click: one LG.</summary>
public sealed record LgDetail
{
#pragma warning disable CS1591
    public string LgNo { get; init; } = "";
    public bool Pending { get; init; }
    public List<string> Messages { get; init; } = [];
    public string Member { get; init; } = "";
    public string Ic { get; init; } = "";
    public string Patient { get; init; } = "";
    public string GlType { get; init; } = "";
    public string DeclineReason { get; init; } = "";
    public string InternalRemarks { get; init; } = "";
    public string Remarks { get; init; } = "";
    public LgLimits Limits { get; init; } = new();
    public bool Pharmacy { get; init; }
    public bool MasterCase { get; init; }
    public bool MedRequestInProgress { get; init; }
    public LgTemplate Template { get; init; } = new();
    public List<string> Documents { get; init; } = [];
#pragma warning restore CS1591
}

/// <summary>ValidateLG1: the clinic and the GL template.</summary>
public sealed record LgTemplate
{
#pragma warning disable CS1591
    public string GlFormat { get; init; } = "";
    public string ClinicName { get; init; } = "";
    public string ClinicAddress1 { get; init; } = "";
    public string ClinicAddress2 { get; init; } = "";
    public string ClinicState { get; init; } = "";
    public string ClinicPanel { get; init; } = "";
    public string TemplateName { get; init; } = "";
    public string DependantTemplate { get; init; } = "";
    public string DependantInd { get; init; } = "";
    public string VipInd { get; init; } = "";
    public string Remarks { get; init; } = "";
#pragma warning restore CS1591
}

/// <summary>Form_Load drop-downs.</summary>
public record LgLookups(List<string> GlTypes, List<string> DeclineReasons, List<string> DelayReasons);

/// <summary>The data Lg_Member returns.</summary>
public sealed record LgMemberData
{
#pragma warning disable CS1591
    public Dictionary<string, object?>? Limits { get; init; }
    public Dictionary<string, object?>? Validate { get; init; }
    public decimal SpTopup { get; init; }
    public string? LimitStatus { get; init; }
    public bool Maternity { get; init; }
    public bool Dental { get; init; }
    public bool PrinDepGpShared { get; init; }
    public (decimal Limit, decimal Utilised)? Aso { get; init; }
#pragma warning restore CS1591
}

/// <summary>The data Lg_Load returns.</summary>
public sealed record LgLoadData(List<Dictionary<string, object?>> Rows, bool Pending, bool Pharmacy, string? MainCase, (string Name, string Ic)? Offline);

/// <summary>Data access of Generate LG.</summary>
public interface ILgRepository
{
#pragma warning disable CS1591
    Task<(List<Dictionary<string, object?>> Rows, string? Warranty)> SearchAsync(IDbSession session, string member);
    Task<LgMemberData> MemberAsync(IDbSession session, string member, string cover, string ic);
    Task<decimal?> MerchantradeAsync(IDbSession session, string member);
    Task<List<Dictionary<string, object?>>> ListAsync(IDbSession session, string member, string? cover);
    Task<LgLoadData> LoadAsync(IDbSession session, string lgNo);
    Task<bool> MedRequestInProgressAsync(IDbSession session, string caseNo);
    Task<LgTemplate> TemplateAsync(IDbSession session, string clinic, string company, string plan, string glType, string ic, string category, string caseNo);
    Task<List<Dictionary<string, object?>>> CasesAsync(IDbSession session, string member, string? cover, bool specialist);
    Task<LgLookups> LookupsAsync(IDbSession session);
    Task<List<(string Code, string Name)>> ClinicsAsync(IDbSession session, string company);
#pragma warning restore CS1591
}

/// <summary>LG &amp; Audit > Generate LG (frmLGGeneration).</summary>
public interface ILgService
{
    /// <summary>cmdSearch_Click.</summary>
    Task<LgSearchResult> SearchAsync(string? member);

    /// <summary>lstMBM_Click.</summary>
    Task<LgMemberInfo> MemberAsync(string? member, string? cover);

    /// <summary>lstLG_Click.</summary>
    Task<LgDetail> LoadAsync(string? lgNo, string? groupCompany, string? category);

    /// <summary>The case tabs.</summary>
    Task<List<Dictionary<string, object?>>> CasesAsync(string? member, string? cover, bool specialist);

    /// <summary>Drop-downs.</summary>
    Task<LgLookups> LookupsAsync();

    /// <summary>Group company clinics.</summary>
    Task<List<(string Code, string Name)>> ClinicsAsync(string? company);

    /// <summary>CmdAdd_Click.</summary>
    Task<LgGateResult> AddCheckAsync(string? member, string? cover, int confirmed);

    /// <summary>CmdEdit_Click.</summary>
    Task<LgGateResult> EditCheckAsync(string? lgNo, string? clinicText, string? category, int confirmed);

    /// <summary>cmdLG_Click + SaveLG.</summary>
    Task<LgActionResult> IssueAsync(LgIssueRequest request, string user);

    /// <summary>cmdDecline_Click.</summary>
    Task<LgActionResult> DeclineAsync(LgDeclineRequest request, string user);

    /// <summary>CmdPreview_Click.</summary>
    Task<Documents.FileOutput> PreviewAsync(LgIssueRequest request, string user);

    /// <summary>cmdCopyFile_Click.</summary>
    Task<string> CopyAsync(string? lgNo, string? fileName, Stream content);

    /// <summary>A document of the LG.</summary>
    Task<(byte[] Content, string FileName)> DocumentAsync(string? lgNo, string? fileName);
}

/// <inheritdoc />
public sealed partial class LgService(IDbSessions sessions, ILgRepository repository, ILgSaveRepository saves, IClinicInvoiceRepository clinics,
    IMbmEnquiryRepository enquiry, IFileDrop files, Documents.ILetterWriter letters, Email.IEmailSender mail, TimeProvider clock) : ILgService
{
    /// <summary>Desktop texts.</summary>
    public const string KeyMember = "Please key the Membership No", NotFound = "Record not found!", NotActive = "Member is not active.",
        NoProviders = "Provider Listing is not available!\n Please verify with Administrator!",
        PendingNotice = "The case is still in pending state.\nIf you proceed to issue GL, the pending status is turned back to 'Accepted' and the selected pending reasons will be cleared.";

    /// <summary>Companies (and members) whose warranty text SearchRecords shows.</summary>
    private static readonly string[] WarrantyCompanies =
        ["OPENET TELECOM MALAYSIA SDN BHD", "SURUHANJAYA PERKHIDMATAN AIR NEGARA", "COMMUNICATION CORPORATION SDN BHD", "SAINT-GOBAIN MALAYSIA SDN BHD",
         "PERBADANAN KEMAJUAN NEGERI SELANGOR", "DIGI TELECOMMUNICATIONS SDN BHD"];

    private static readonly string[] WarrantyMembers = ["STHA3342413*01", "STHA3342413*02", "STHB3164093*01"];

    /// <summary>cmdSearch_Click / SearchRecords.</summary>
    public async Task<LgSearchResult> SearchAsync(string? member)
    {
        var number = Member(member);
        await using var s = await sessions.OpenAsync();
        var (rows, warranty) = await repository.SearchAsync(s, number);
        if (rows.Count == 0) throw new NotFoundException(NotFound);
        var list = rows.Select(Row).ToList();
        var group = list[^1].GroupCompany;
        var show = group != "INTERNATIONAL RESCUE COMMITTEE1" && (WarrantyCompanies.Contains(group) || WarrantyMembers.Contains(number));
        return new LgSearchResult(list, show && !string.IsNullOrWhiteSpace(warranty) ? warranty.Trim() : null);
    }

    /// <summary>lstMBM_Click: CheckMBMLimit, ValidateLG and LGListing for one cover.</summary>
    public async Task<LgMemberInfo> MemberAsync(string? member, string? cover)
    {
        var number = Member(member);
        var cov = Cover(cover);
        await using var s = await sessions.OpenAsync();
        var (rows, _) = await repository.SearchAsync(s, number);
        var all = rows.Select(Row).ToList();
        var row = all.FirstOrDefault(r => r.Cover == cov) ?? throw new NotFoundException(NotFound);
        var d = await repository.MemberAsync(s, number, cov, row.Ic);
        var lim = d.Limits ?? throw new NotFoundException(NotFound);
        var specLimit = "";
        var messages = new List<string>();
        if (!row.Aso)
        {
            var spec = await enquiry.SpecLimitAsync(s, T(lim, "PLNCode"), "DC", "H");
            specLimit = MbmEnquiryService.SpecLimit(spec, [], cov != "00");
        }
        var input = new LgLimitInput
        {
            CoverId = cov, Row = lim, Aso = row.Aso, AsoValues = d.Aso, SpTopup = d.SpTopup, LimitStatus = d.LimitStatus, SpecLimit = specLimit,
            Maternity = d.Maternity, Dental = d.Dental, PrinDepGpShared = d.PrinDepGpShared,
        };
        var limitStatus = d.LimitStatus is null ? "" : d.LimitStatus.Trim() is { Length: > 0 } ls ? ls : "A";
        var mode = LgRules.ClaimMode(limitStatus, cov, d.PrinDepGpShared, LgRules.Fmt(T(lim, "OriSpecAnnLimit")) != "N/A", T(lim, "SpecSuppLimitStatus"));
        var claims = mode is { } m ? await clinics.ClaimsAsync(s, number, m, cov) : [];
        var limits = LgRules.Limits(input, claims);
        if (row.GroupCompany == "MERCHANTRADE ASIA SDN BHD" && LgRules.IsNum(limits.AnnualLimit) && limits.PlanCode == "M230"
            && await repository.MerchantradeAsync(s, number) is { } m230)
        {
            limits = limits with { Balance = LgRules.Fmt(m230.ToString(CultureInfo.InvariantCulture)) };
            messages.Add($"Current Month Balance Limit For Plan Code M230. Amount RM {limits.Balance} ");
        }
        var v = d.Validate;
        var gpSp = T(v, "PLNGPSPShared") is { Length: > 0 } g ? g : "N";
        var after = T(v, "LGIssuedAfterLimitExInd") is { Length: > 0 } a ? a : "N";
        if (v is not null && T(v, "ANNLmtIND") == "N") limits = limits with { AnnualLimit = "FULL" };
        var memberStatus = all[^1].ClinicStatus == "C" ? "C" : T(v, "MBMStatus");
        var valid = LgRules.IsValidLg(T(v, "LGIssuedInd"), gpSp, after, limits.Balance, limits.SpecBalance);
        if (memberStatus != "A") messages.Add(NotActive);
        var lgs = await repository.ListAsync(s, number, cov);
        return new LgMemberInfo(row, limits, memberStatus, valid, gpSp, after, messages, lgs);
    }

    /// <summary>lstLG_Click: one LG with its limits, template, pharmacy state and documents.</summary>
    public async Task<LgDetail> LoadAsync(string? lgNo, string? groupCompany, string? category)
    {
        var lg = Case(lgNo);
        await using var s = await sessions.OpenAsync();
        var d = await repository.LoadAsync(s, lg);
        var row = d.Rows.FirstOrDefault(r => T(r, "CoverID") == T(r, "MBMCOVID")) ?? d.Rows.FirstOrDefault(r => T(r, "CoverID") == T(r, "MBMCCoverID"))
            ?? throw new NotFoundException(NotFound);
        var principal = T(row, "CoverID") == T(row, "MBMCOVID");
        var member = T(row, "MBMNumber");
        var cover = principal ? T(row, "MBMCOVID") : T(row, "MBMCCoverID");
        var list = await repository.ListAsync(s, member, cover);
        var item = list.FirstOrDefault(r => T(r, "LGNo") == lg) ?? throw new NotFoundException(NotFound);
        var lim = (await repository.MemberAsync(s, member, cover, principal ? T(row, "MBMIcBcPp") : T(row, "MBMCICBCPPNo"))).Limits ?? [];
        var messages = new List<string>();
        if (d.Pending) messages.Add(PendingNotice);
        if (T(item, "Alertmsg") is { Length: > 0 } alert) messages.Add(alert);
        if (T(lim, "Error") is { Length: > 0 } err && err != "N/A") messages.Add(err);
        var ic = d.Offline?.Ic.Replace("-", "").Trim() ?? "";
        var template = await repository.TemplateAsync(s, T(item, "CLNName"), (groupCompany ?? "").Trim(), T(lim, "PLNCode"), T(item, "GLType"), ic,
            (category ?? "").Trim(), lg);
        var inProgress = d.Pharmacy && await repository.MedRequestInProgressAsync(s, d.MainCase ?? lg);
        return new LgDetail
        {
            LgNo = lg, Pending = d.Pending, Messages = messages, Member = $"{member}-{cover}",
            Ic = principal ? T(row, "MBMIcBcPp") : T(row, "MBMCICBCPPNo"), Patient = principal ? T(row, "MBMName") : T(row, "MBMCName"),
            GlType = T(row, "GLType"), DeclineReason = T(row, "DeclineReason"), InternalRemarks = T(item, "InternalRemarks"),
            Remarks = T(item, "LGRemarks") is { Length: > 0 } rm ? rm : template.Remarks, Limits = LgRules.RawLimits(lim, null, false),
            Pharmacy = d.Pharmacy, MasterCase = !string.IsNullOrWhiteSpace(d.MainCase), MedRequestInProgress = inProgress, Template = template,
            Documents = Documents(lg, d.Pharmacy),
        };
    }

    /// <summary>ListCases / ListSpecialistCases.</summary>
    public async Task<List<Dictionary<string, object?>>> CasesAsync(string? member, string? cover, bool specialist)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.CasesAsync(s, Member(member), string.IsNullOrWhiteSpace(cover) ? null : Cover(cover), specialist);
    }

    /// <summary>Form_Load drop-downs.</summary>
    public async Task<LgLookups> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s);
    }

    /// <summary>chkCLN_Click: the group company's clinics.</summary>
    public async Task<List<(string Code, string Name)>> ClinicsAsync(string? company)
    {
        if (string.IsNullOrWhiteSpace(company)) return [];
        await using var s = await sessions.OpenAsync();
        var list = await repository.ClinicsAsync(s, company.Trim());
        return list.Count == 0 ? throw new BusinessRuleException(NoProviders) : list;
    }

    /// <summary>addTheFilesClaims: the LG's documents (pharmacy cases: OtherScannedDoc).</summary>
    private List<string> Documents(string lg, bool pharmacy)
    {
        if (Folder(lg, pharmacy) is not { } dir || !Directory.Exists(dir)) return [];
        string[] types = [".jpg", ".tif", ".gif", ".pdf", ".bmp", ".png", ".msg"];
        return Directory.EnumerateFiles(dir).Select(Path.GetFileName).OfType<string>()
            .Where(f => types.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase)).Order(StringComparer.OrdinalIgnoreCase).ToList();
    }

    private string? Folder(string lg, bool pharmacy) =>
        files.FolderOf(FileFolders.Outpatient) is { Length: > 0 } root ? Path.Combine(root, pharmacy ? "OtherScannedDoc" : "Guarantee Letter", lg) : null;

    private static LgMemberRow Row(Dictionary<string, object?> r) => new()
    {
        Member = T(r, "MBMNumber"), Cover = T(r, "MBMCOVID"), Name = T(r, "MBMName"), Ic = T(r, "MBMIcBcPp"), GroupCompany = T(r, "MBMGroupCompany"),
        StaffCategory = T(r, "MBMStaffCategory"), ClinicStatus = T(r, "MBMCLNStatus"), Plan = T(r, "PLNType"), PlanLgRemarks = T(r, "PLNLGRemarks"),
        PolicyExpiry = r.GetValueOrDefault("MBMPolExpDate") as DateTime?, PlanSpecialRemark = T(r, "PLNSpecialRemark"), Aso = T(r, "ASOInd") == "Y",
        ClinicalRemarks = T(r, "clinicalRemarks"), Category = T(r, "MBMCategory"), Subsidiary = T(r, "Subsidiary"),
    };

    /// <summary>A member number (the part before "-").</summary>
    internal static string Member(string? member)
    {
        var m = (member ?? "").Trim();
        if (m.IndexOf('-') is > 0 and var i) m = m[..i].Trim();
        if (m.Length == 0) throw new BusinessRuleException(KeyMember);
        return m.Length <= 30 && m.All(c => char.IsLetterOrDigit(c) || c == '*') ? m.ToUpperInvariant() : throw new NotFoundException(NotFound);
    }

    private static string Cover(string? cover)
    {
        var c = (cover ?? "").Trim();
        return c.Length is > 0 and <= 2 && c.All(char.IsDigit) ? c : throw new NotFoundException(NotFound);
    }

    internal static string Case(string? lgNo)
    {
        var c = (lgNo ?? "").Trim().ToUpperInvariant();
        return c.Length is > 0 and <= 20 && c.All(char.IsLetterOrDigit) ? c : throw new NotFoundException(NotFound);
    }

    internal static string T(IReadOnlyDictionary<string, object?>? r, string k) =>
        r is not null && r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";
}
