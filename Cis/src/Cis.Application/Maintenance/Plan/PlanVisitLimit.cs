using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.Plan;

/// <summary>Visit Limit tab fields (frmMaintenancePlan tab 5). Combos carry their code (or "CODE - NAME").</summary>
/// <param name="Index">VisitIndex of the selected row (Edit / Delete).</param>
/// <param name="PlnCode">Plan.</param>
/// <param name="HltCode">Health Type.</param>
/// <param name="InsCode">Insured Type.</param>
/// <param name="PayCode">Payor (search only).</param>
/// <param name="GrpCompany">Group company (search only, prefix).</param>
/// <param name="VisitLimit">Visit limit.</param>
/// <param name="SuppStatus">A / B / C (Supp. Visit Lmt Status).</param>
/// <param name="SuppLimit">Supp. Visit Lmt (B / C).</param>
/// <param name="EffDate">Effective date dd/mm/yyyy.</param>
/// <param name="Version">Visit Limit Version.</param>
public record VisitLimitDto(int? Index, string? PlnCode, string? HltCode, string? InsCode, string? PayCode, string? GrpCompany, string? VisitLimit,
    string? SuppStatus, string? SuppLimit, string? EffDate, string? Version);

/// <summary>Save / search of the Visit Limit tab.</summary>
/// <param name="Limit">The fields.</param>
public record VisitLimitRequest(VisitLimitDto Limit);

/// <summary>A VisitLimit row as SaveVisit writes it.</summary>
/// <param name="Index">VisitIndex (null = Add).</param>
/// <param name="PlnCode">PLNCode.</param>
/// <param name="InsCode">INSCode.</param>
/// <param name="HltCode">HLTCode.</param>
/// <param name="VisitLimit">VisitLimit.</param>
/// <param name="Status">SuppVisitLimitStatus.</param>
/// <param name="SuppLimit">SuppVisitLimit.</param>
/// <param name="EffDate">VisitEffDate.</param>
/// <param name="Version">VisitLimitVersion.</param>
public record VisitLimitRow(int? Index, string PlnCode, string InsCode, string HltCode, decimal VisitLimit, string Status, decimal? SuppLimit,
    DateTime EffDate, int Version);

/// <summary>Number formats.</summary>
public sealed class VisitLimitRequestValidator : RequestValidator<VisitLimitRequest>
{
    /// <summary>Defines the rules.</summary>
    public VisitLimitRequestValidator() => RuleFor(r => r.Limit).NotNull().ChildRules(l =>
    {
        l.RuleFor(x => x.VisitLimit).Must(v => string.IsNullOrWhiteSpace(v) || PlanAnnualLimitService.Money(v) is not null).WithMessage("Visit Limit must be a number.");
        l.RuleFor(x => x.SuppLimit).Must(v => string.IsNullOrWhiteSpace(v) || PlanAnnualLimitService.Money(v) is not null).WithMessage("Supp Visit Limit must be a number.");
        l.RuleFor(x => x.Version).Must(v => string.IsNullOrWhiteSpace(v) || int.TryParse(v.Trim(), out _)).WithMessage("Visit Limit Version must be a whole number.");
    });
}

/// <summary>Data access of the Visit Limit tab.</summary>
public interface IPlanVisitLimitRepository
{
    /// <summary>VIEW_VISIT_PAY_PLN rows.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, string? pln, string? hlt, string? ins, string? pay, string? grp);

    /// <summary>True when insured + health + plan + version + date exists.</summary>
    Task<bool> ExistsAsync(IDbSession session, VisitLimitRow row);

    /// <summary>Insert / update; rows written.</summary>
    Task<int> SaveAsync(IDbSession session, VisitLimitRow row, string userCode, DateTime today);

    /// <summary>Delete; rows removed.</summary>
    Task<int> DeleteAsync(IDbSession session, int index);
}

/// <summary>Maintenance > Plan, Visit Limit tab (combos: the Annual Limit setup).</summary>
public interface IPlanVisitLimitService
{
    /// <summary>SearchVisit; "Record not found" when empty.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(VisitLimitDto filters);

    /// <summary>SaveVisit.</summary>
    Task<string> SaveAsync(VisitLimitDto limit, bool editing, string userCode);

    /// <summary>DeleteVisit.</summary>
    Task<string> DeleteAsync(VisitLimitDto limit);
}

/// <inheritdoc />
public sealed class PlanVisitLimitService(IDbSessions sessions, IPlanVisitLimitRepository repository, IPlanAnnualLimitRepository annual, TimeProvider clock)
    : IPlanVisitLimitService
{
    /// <summary>SaveVisit's checks in order.</summary>
    public static readonly IReadOnlyList<(Func<VisitLimitDto, bool> Missing, string Message)> SaveChecks =
    [
        (v => E(v.PlnCode), "Please Enter Plan Code"),
        (v => E(v.HltCode), "Please Enter Health Code"),
        (v => E(v.InsCode), "Please Enter Insured Code"),
        (v => E(v.VisitLimit), "Please Enter Visit Limit"),
        (v => E(v.EffDate) || v.EffDate!.Trim() == "__/__/____", "Please Enter Visit Limit Effective Date"),
        (v => E(v.Version), "Please Enter Visit Limit Version"),
        (v => Status(v) is "B" or "C" && E(v.SuppLimit), "Please Enter Supp Visit Limit"),
    ];

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(VisitLimitDto f)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, Opt(Before(f.PlnCode, 4)), Opt(Before(f.HltCode, 4)), Opt(Before(f.InsCode, 4)),
            Opt(Before(f.PayCode, 4)), Opt(Before(f.GrpCompany, 4)));
        return rows.Count == 0 ? throw new NotFoundException(PlanAnnualLimitService.NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(VisitLimitDto v, bool editing, string userCode)
    {
        foreach (var (missing, message) in SaveChecks)
            if (missing(v)) throw new BusinessRuleException(message);
        if (editing && v.Index is null) throw new BusinessRuleException(PlanAnnualLimitService.SearchFirst);
        var row = Row(v, editing);
        await using var s = await sessions.BeginAsync();
        var plan4 = (v.PlnCode ?? "").Length > 4 ? v.PlnCode![..4].Trim() : (v.PlnCode ?? "").Trim();
        if (!editing && await repository.ExistsAsync(s, row with { PlnCode = plan4 })) throw new ConflictException(PlanAnnualLimitService.Duplicated);
        if (await repository.SaveAsync(s, row, userCode, clock.GetLocalNow().Date) == 0) throw new NotFoundException(PlanAnnualLimitService.SearchFirst);
        await s.CommitAsync();
        return editing ? PlanAnnualLimitService.Updated : PlanAnnualLimitService.Saved;
    }

    /// <inheritdoc />
    public async Task<string> DeleteAsync(VisitLimitDto v)
    {
        if (string.IsNullOrEmpty(v.PlnCode)) throw new BusinessRuleException(PlanAnnualLimitService.PlanFirst);
        if (v.Index is not { } index) throw new BusinessRuleException(PlanAnnualLimitService.SearchFirst);
        await using var s = await sessions.BeginAsync();
        if (await annual.InUseAsync(s, v.PlnCode.Trim())) throw new ConflictException(PlanAnnualLimitService.InUse);
        if (await repository.DeleteAsync(s, index) == 0) throw new NotFoundException(PlanAnnualLimitService.SearchFirst);
        await s.CommitAsync();
        return PlanAnnualLimitService.Deleted;
    }

    /// <summary>SaveVisit's values: plan before "-" (else first 4), insured / health first character, limit, status, supplementary limit
    /// only for B / C, date and version.</summary>
    public static VisitLimitRow Row(VisitLimitDto v, bool editing)
    {
        var status = Status(v) is { Length: 1 } st ? st : "A";
        var eff = DateTime.TryParseExact(v.EffDate?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d)
            ? d : throw new BusinessRuleException(PlanAnnualLimitService.BadDate);
        return new(editing ? v.Index : null, Before(v.PlnCode, 4), First(v.InsCode), First(v.HltCode), PlanAnnualLimitService.Money(v.VisitLimit) ?? 0,
            status, status == "A" ? null : PlanAnnualLimitService.Money(v.SuppLimit), eff, int.Parse(v.Version!.Trim(), CultureInfo.InvariantCulture));
    }

    private static string Status(VisitLimitDto v) => string.IsNullOrWhiteSpace(v.SuppStatus) ? "" : v.SuppStatus.Trim()[..1];

    private static string First(string? v) => string.IsNullOrEmpty(v) ? "" : v[..1];

    private static bool E(string? v) => string.IsNullOrWhiteSpace(v);

    private static string? Opt(string v) => v.Length == 0 ? null : v;

    private static string Before(string? text, int n)
    {
        if (E(text)) return "";
        var at = text!.IndexOf('-');
        return (at < 0 ? text[..Math.Min(n, text.Length)] : text[..at]).Trim();
    }
}
