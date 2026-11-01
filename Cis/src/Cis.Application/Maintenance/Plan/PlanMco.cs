using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.Plan;

/// <summary>MCO Fee tab fields (frmMaintenancePlan tab 6). Fees by insured type: I / G adult + child, F / H adult + child + supplementary.</summary>
/// <param name="HltCode">Health Type.</param>
/// <param name="PlnCode">Plan.</param>
/// <param name="SuppStatus">A / B / C (Supp MCO status).</param>
/// <param name="EffDate">Effective date dd/mm/yyyy.</param>
/// <param name="Fees">Fee boxes by name: IA, IC, FA, FC, SFA, SFC, GA, GC, HA, HC, SHA, SHC.</param>
public record McoDto(string? HltCode, string? PlnCode, string? SuppStatus, string? EffDate, Dictionary<string, string?>? Fees);

/// <summary>Save / search of the MCO Fee tab.</summary>
/// <param name="Mco">The fields.</param>
public record McoRequest(McoDto Mco);

/// <summary>An MCO row as SaveMCO1 writes it.</summary>
/// <param name="InsCode">Insured type.</param>
/// <param name="Adult">MCOAmtAdult.</param>
/// <param name="Child">MCOAmtChild.</param>
/// <param name="SuppAdult">MCOSuppAdultAmt.</param>
/// <param name="SuppChild">MCOSuppChildAmt.</param>
public record McoRow(string InsCode, double Adult, double Child, double SuppAdult, double SuppChild);

/// <summary>Combos of the MCO Fee tab.</summary>
/// <param name="HealthTypes">HLT.</param>
/// <param name="Plans">PLN codes.</param>
public record McoSetup(List<LookupOption> HealthTypes, List<LookupOption> Plans);

/// <summary>Fee box formats.</summary>
public sealed class McoRequestValidator : RequestValidator<McoRequest>
{
    /// <summary>Defines the rules.</summary>
    public McoRequestValidator() => RuleFor(r => r.Mco).NotNull();
}

/// <summary>Data access of the MCO Fee tab.</summary>
public interface IPlanMcoRepository
{
    /// <summary>Combos.</summary>
    Task<McoSetup> SetupAsync(IDbSession session);

    /// <summary>MCO rows.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, string? hlt, string? pln, DateTime? eff);

    /// <summary>True when plan + health + date + insured type exists.</summary>
    Task<bool> ExistsAsync(IDbSession session, string pln, string hlt, DateTime eff, string ins);

    /// <summary>Insert one insured type.</summary>
    Task InsertAsync(IDbSession session, string pln, string hlt, string supp, DateTime eff, McoRow row, string userCode, DateTime today);

    /// <summary>Delete one row; rows removed.</summary>
    Task<int> DeleteAsync(IDbSession session, int index);
}

/// <summary>Maintenance > Plan, MCO Fee tab.</summary>
public interface IPlanMcoService
{
    /// <summary>MCOComboListing.</summary>
    Task<McoSetup> SetupAsync();

    /// <summary>SearchMCO; "Record not found" when empty.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(McoDto filters);

    /// <summary>SaveMCO: the per-type refusals, then "Record Saved ! ".</summary>
    Task<List<string>> SaveAsync(McoDto mco, string userCode);

    /// <summary>DeleteMCO (the selected row).</summary>
    Task<string> DeleteAsync(int index, string? hltCode);
}

/// <inheritdoc />
public sealed class PlanMcoService(IDbSessions sessions, IPlanMcoRepository repository, TimeProvider clock) : IPlanMcoService
{
    /// <summary>Desktop texts (spacing kept).</summary>
    public const string HltFirst = "Please Enter Health Code", PlnFirst = "Please Enter Plan Code", FeesFirst = "Please Enter MCO Fees ",
        BadDate = "Please validate your MCO Effective date", Saved = "Record Saved ! ", NotFound = "Record not found",
        HltDelete = "Please Enter Health Type", SearchFirst = "Please search the record", Deleted = "Record Deleted...";

    private static readonly string[] Principal = ["IA", "IC", "FA", "FC", "GA", "GC", "HA", "HC"];

    /// <inheritdoc />
    public async Task<McoSetup> SetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.SetupAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(McoDto f)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, E(f.HltCode) ? null : f.HltCode!.Trim()[..1], E(f.PlnCode) ? null : f.PlnCode!.Trim(), ParseDate(f.EffDate));
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<List<string>> SaveAsync(McoDto m, string userCode)
    {
        if (E(m.HltCode)) throw new BusinessRuleException(HltFirst);
        if (E(m.PlnCode)) throw new BusinessRuleException(PlnFirst);
        if (Principal.All(k => E(Fee(m, k)))) throw new BusinessRuleException(FeesFirst);
        var eff = ParseDate(m.EffDate) ?? throw new BusinessRuleException(BadDate);
        var (pln, hlt) = (m.PlnCode!.Trim(), m.HltCode!.Trim()[..1]);
        var supp = E(m.SuppStatus) ? "" : m.SuppStatus!.Trim()[..1];
        var messages = new List<string>();
        await using var s = await sessions.BeginAsync();
        var today = clock.GetLocalNow().Date;
        foreach (var row in Rows(m))
        {
            if (await repository.ExistsAsync(s, pln, hlt, eff, row.InsCode))
                messages.Add($"Record Not Save for Insured Type '{row.InsCode}! Please Change The Corresponding Field");
            else
                await repository.InsertAsync(s, pln, hlt, supp, eff, row, userCode, today);
        }
        await s.CommitAsync();
        messages.Add(Saved);
        return messages;
    }

    /// <inheritdoc />
    public async Task<string> DeleteAsync(int index, string? hltCode)
    {
        if (E(hltCode)) throw new BusinessRuleException(HltDelete);
        if (index <= 0) throw new BusinessRuleException(SearchFirst);
        await using var s = await sessions.BeginAsync();
        if (await repository.DeleteAsync(s, index) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return Deleted;
    }

    /// <summary>SaveMCO's rows: I and G adult / child (no supplementary), F and H with supplementary; non-numeric boxes are 0; a type is
    /// written only when one of its amounts is above 0.</summary>
    public static List<McoRow> Rows(McoDto m)
    {
        double N(string k) => double.TryParse(Fee(m, k)?.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : 0;
        McoRow[] all =
        [
            new("I", N("IA"), N("IC"), 0, 0), new("F", N("FA"), N("FC"), N("SFA"), N("SFC")),
            new("G", N("GA"), N("GC"), 0, 0), new("H", N("HA"), N("HC"), N("SHA"), N("SHC")),
        ];
        return all.Where(r => r.Adult > 0 || r.Child > 0 || r.SuppAdult > 0 || r.SuppChild > 0).ToList();
    }

    /// <summary>The masked dd/mm/yyyy box (IsDate); ISO accepted too.</summary>
    public static DateTime? ParseDate(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static string? Fee(McoDto m, string key) => m.Fees is not null && m.Fees.TryGetValue(key, out var v) ? v : null;

    private static bool E(string? v) => string.IsNullOrWhiteSpace(v);
}
