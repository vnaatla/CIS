using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.LgTemplates;

/// <summary>A Tbl_OPGLMaster row (lstLGTemplate).</summary>
/// <param name="Id">ID.</param>
/// <param name="GroupCompany">GRPCompnay (group company name).</param>
/// <param name="PayCode">Paycode.</param>
/// <param name="PlanCode">PLNCode.</param>
/// <param name="GlType">GLType (name).</param>
/// <param name="GlTemplate">GLTemplate.</param>
/// <param name="LgFormat">LGFormat.</param>
/// <param name="VipInd">VIPInd (Y / N).</param>
/// <param name="DepInd">DepInd (Y / N).</param>
/// <param name="DepTemplate">DepTemplate.</param>
/// <param name="TmpRemarks">TmpRemarks.</param>
public record LgBinding(int Id, string? GroupCompany, string? PayCode, string? PlanCode, string? GlType, string? GlTemplate, string? LgFormat,
    string? VipInd, string? DepInd, string? DepTemplate, string? TmpRemarks);

/// <summary>The tab's drop-downs.</summary>
/// <param name="GroupCompanies">GRPCompany (code, name).</param>
/// <param name="GlTypes">Tbl_GLType (GLID, GLType).</param>
/// <param name="Templates">Tbl_GLTemplates (name, remarks).</param>
/// <param name="LgFormats">HISDB.CLN.CLNGLFormat values.</param>
public record LgBindingLookups(List<LookupOption> GroupCompanies, List<LookupOption> GlTypes, List<LookupOption> Templates, List<LookupOption> LgFormats);

/// <summary>Plans and pay codes of a group company (BindPlanCode / BindPayCode).</summary>
/// <param name="Plans">PLN (code, description).</param>
/// <param name="PayCodes">Distinct PLN.PAYCode.</param>
public record LgGroupPlans(List<LookupOption> Plans, List<LookupOption> PayCodes);

/// <summary>Save of the tab (cmdSave_Click with SSTab1.Tab = 0).</summary>
/// <param name="GroupName">Group company name (the desktop stored the part after "-" of "CODE - NAME").</param>
/// <param name="PlanCode">Plan code.</param>
/// <param name="GlType">GL type name.</param>
/// <param name="Template">Template name.</param>
/// <param name="PayCode">Pay code.</param>
/// <param name="LgFormat">LG format.</param>
/// <param name="Vip">"Y" or "N".</param>
/// <param name="DependentTemplate">Dependent template (empty = DepInd N).</param>
/// <param name="Remarks">Template remarks.</param>
public record LgBindingSaveRequest(string? GroupName, string? PlanCode, string? GlType, string? Template, string? PayCode, string? LgFormat, string? Vip,
    string? DependentTemplate, string? Remarks);

/// <summary>Lengths of Tbl_OPGLMaster; the desktop had no other check.</summary>
public sealed class LgBindingSaveRequestValidator : RequestValidator<LgBindingSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public LgBindingSaveRequestValidator()
    {
        RuleFor(r => r.PlanCode).MaximumLength(4).WithMessage("Plan Code must be at most 4 characters.");
        RuleFor(r => r.LgFormat).MaximumLength(5).WithMessage("LG Format must be at most 5 characters.");
        RuleFor(r => r.Vip).Must(v => string.IsNullOrEmpty(v) || v is "Y" or "N").WithMessage("VIP Indicator must be Y or N.");
    }
}

/// <summary>Data access of the GL Template Binding tab.</summary>
public interface ILgBindingRepository
{
    /// <summary>Drop-downs.</summary>
    Task<LgBindingLookups> LookupsAsync(IDbSession session);

    /// <summary>Plans and pay codes of the group company code.</summary>
    Task<LgGroupPlans> PlansOfGroupAsync(IDbSession session, string groupCode);

    /// <summary>Rows of the group company name (LIKE), optional pay code / plan code.</summary>
    Task<List<LgBinding>> SearchAsync(IDbSession session, string groupName, string? payCode, string? planCode);

    /// <summary>Insert.</summary>
    Task InsertAsync(IDbSession session, LgBindingSaveRequest request, string depInd);

    /// <summary>Update every row of the plan; returns the rows changed.</summary>
    Task<int> UpdatePlanAsync(IDbSession session, LgBindingSaveRequest request, string depInd);
}

/// <summary>GL Template Binding with the desktop's behaviour and messages.</summary>
public interface ILgBindingService
{
    /// <summary>Drop-downs (Form_Load / CmdClear_Click).</summary>
    Task<LgBindingLookups> LookupsAsync();

    /// <summary>Plans and pay codes of a group company (cmbGrpCompany_LostFocus).</summary>
    Task<LgGroupPlans> PlansOfGroupAsync(string groupCode);

    /// <summary>SearchGLTemplate: group company required ("Plese Select Group Company to Search").</summary>
    Task<List<LgBinding>> SearchAsync(string? groupName, string? payCode, string? planCode);

    /// <summary>Add + Save → "New Record Inserted Successfully..".</summary>
    Task<string> CreateAsync(LgBindingSaveRequest request);

    /// <summary>Edit + Save → "Record Updated Successfully.." (updates every row of the plan code, as the desktop did).</summary>
    Task<string> UpdateAsync(LgBindingSaveRequest request);
}

/// <inheritdoc />
public sealed class LgBindingService(IDbSessions sessions, ILgBindingRepository repository) : ILgBindingService
{
    /// <summary>Desktop texts (spelling kept).</summary>
    public const string SelectGroup = "Plese Select Group Company to Search", Inserted = "New Record Inserted Successfully..",
        Updated = "Record Updated Successfully..";

    /// <inheritdoc />
    public async Task<LgBindingLookups> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s);
    }

    /// <inheritdoc />
    public async Task<LgGroupPlans> PlansOfGroupAsync(string groupCode)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.PlansOfGroupAsync(s, groupCode.Trim());
    }

    /// <inheritdoc />
    public async Task<List<LgBinding>> SearchAsync(string? groupName, string? payCode, string? planCode)
    {
        if (string.IsNullOrWhiteSpace(groupName)) throw new BusinessRuleException(SelectGroup);
        await using var s = await sessions.OpenAsync();
        return await repository.SearchAsync(s, "%" + groupName.Trim() + "%", Blank(payCode), Blank(planCode));
    }

    /// <inheritdoc />
    public async Task<string> CreateAsync(LgBindingSaveRequest request)
    {
        await using var s = await sessions.BeginAsync();
        await repository.InsertAsync(s, Clean(request), DepInd(request));
        await s.CommitAsync();
        return Inserted;
    }

    /// <inheritdoc />
    public async Task<string> UpdateAsync(LgBindingSaveRequest request)
    {
        if (string.IsNullOrWhiteSpace(request.PlanCode)) throw new BusinessRuleException("Please search and select the record to edit.");
        await using var s = await sessions.BeginAsync();
        await repository.UpdatePlanAsync(s, Clean(request), DepInd(request));
        await s.CommitAsync();
        return Updated;
    }

    /// <summary>"Y" when a dependent template is chosen, else "N" (tmpdepind).</summary>
    public static string DepInd(LgBindingSaveRequest r) => string.IsNullOrWhiteSpace(r.DependentTemplate) ? "N" : "Y";

    private static LgBindingSaveRequest Clean(LgBindingSaveRequest r) => new(T(r.GroupName), T(r.PlanCode), T(r.GlType), T(r.Template), T(r.PayCode),
        T(r.LgFormat), T(r.Vip), T(r.DependentTemplate), T(r.Remarks));

    private static string T(string? v) => v?.Trim() ?? "";

    private static string? Blank(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
}
