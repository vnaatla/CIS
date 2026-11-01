using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.Plan;

/// <summary>Premium tab fields (frmMaintenancePlan tab 3). Combos carry their code (or "CODE - NAME").</summary>
/// <param name="InsCode">Insured Type.</param>
/// <param name="HltCode">Health Type.</param>
/// <param name="PayCode">Payor.</param>
/// <param name="AgeCode">Age (used for band rows left empty).</param>
/// <param name="PlnCode">Plan.</param>
/// <param name="EffDate">Effective date dd/mm/yyyy.</param>
/// <param name="Version">Version (whole number).</param>
/// <param name="SuppStatus">A - Share Limit for Family · B - Share Limit for Supplementary · C - Separate Limit for Supplementary.</param>
/// <param name="Amounts">Principal premium of age bands 01–07 (txt01–txt07).</param>
/// <param name="SuppAmounts">Supplementary premium of age bands 01–07 (Text2–Text7, Text1).</param>
public record PremiumDto(string? InsCode, string? HltCode, string? PayCode, string? AgeCode, string? PlnCode, string? EffDate, string? Version,
    string? SuppStatus, List<string?>? Amounts, List<string?>? SuppAmounts);

/// <summary>Save / search of the Premium tab.</summary>
/// <param name="Premium">The fields.</param>
public record PremiumRequest(PremiumDto Premium);

/// <summary>A PRM row as SavePRM writes it.</summary>
/// <param name="InsCode">INSCode.</param>
/// <param name="PayCode">PAYCode.</param>
/// <param name="HltCode">HLTCode.</param>
/// <param name="AgeCode">AGECode.</param>
/// <param name="PlnCode">PLNCode.</param>
/// <param name="Amount">PRMAmount.</param>
/// <param name="SuppAmount">SuppPRMAmt.</param>
/// <param name="SuppStatus">SuppPRMStatus.</param>
/// <param name="EffDate">PRMEffDate.</param>
/// <param name="Version">PRMVersion.</param>
public record PremiumRow(string InsCode, string PayCode, string HltCode, string AgeCode, string PlnCode, decimal? Amount, decimal? SuppAmount,
    string SuppStatus, DateTime EffDate, int? Version);

/// <summary>A band amount of a plan (row click).</summary>
/// <param name="AgeCode">Age code.</param>
/// <param name="Amount">Principal premium.</param>
/// <param name="SuppAmount">Supplementary premium.</param>
public record PremiumBand(string AgeCode, decimal? Amount, decimal? SuppAmount);

/// <summary>Combos of the Premium tab.</summary>
/// <param name="InsuredTypes">INS.</param>
/// <param name="HealthTypes">HLT.</param>
/// <param name="Payors">PAY.</param>
/// <param name="Ages">AGE.</param>
/// <param name="Plans">PLN.</param>
public record PremiumSetup(List<LookupOption> InsuredTypes, List<LookupOption> HealthTypes, List<LookupOption> Payors, List<LookupOption> Ages,
    List<LookupOption> Plans);

/// <summary>Amount and version formats.</summary>
public sealed class PremiumRequestValidator : RequestValidator<PremiumRequest>
{
    /// <summary>Defines the rules.</summary>
    public PremiumRequestValidator() => RuleFor(r => r.Premium).NotNull().ChildRules(p =>
    {
        p.RuleForEach(x => x.Amounts).Must(PlanPremiumService.IsAmountOrEmpty).WithMessage("Premium amounts must be numbers.");
        p.RuleForEach(x => x.SuppAmounts).Must(PlanPremiumService.IsAmountOrEmpty).WithMessage("Supplementary amounts must be numbers.");
        p.RuleFor(x => x.Version).Must(v => string.IsNullOrWhiteSpace(v) || int.TryParse(v.Trim(), out _)).WithMessage("Version must be a whole number.");
    });
}

/// <summary>Data access of the Premium tab.</summary>
public interface IPlanPremiumRepository
{
    /// <summary>Combos.</summary>
    Task<PremiumSetup> SetupAsync(IDbSession session);

    /// <summary>PRM rows.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, string? ins, string? pay, string? hlt, string? age, string? pln);

    /// <summary>Band amounts of insured type + plan.</summary>
    Task<List<PremiumBand>> BandsAsync(IDbSession session, string ins, string pln);

    /// <summary>True when the keys exist.</summary>
    Task<bool> ExistsAsync(IDbSession session, PremiumRow key, string? payCode);

    /// <summary>Insert one row.</summary>
    Task InsertAsync(IDbSession session, PremiumRow row, string userCode, DateTime today);

    /// <summary>True when members use the keys (false where MBM does not exist).</summary>
    Task<bool> InUseAsync(IDbSession session, string ins, string hlt, string pln, string age, string? pay);

    /// <summary>Delete one row; rows removed.</summary>
    Task<int> DeleteAsync(IDbSession session, int prmCode);
}

/// <summary>Maintenance > Plan, Premium tab.</summary>
public interface IPlanPremiumService
{
    /// <summary>PRMComboListing.</summary>
    Task<PremiumSetup> SetupAsync();

    /// <summary>SearchPRM; "Record not found" when empty.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(PremiumDto filters);

    /// <summary>lstPMPRMList_Click band amounts.</summary>
    Task<List<PremiumBand>> BandsAsync(string? ins, string? pln);

    /// <summary>SavePRM (Add).</summary>
    Task<string> SaveAsync(PremiumDto premium, string userCode);

    /// <summary>DeletePRM.</summary>
    Task<string> DeleteAsync(int prmCode, PremiumDto keys);
}

/// <inheritdoc />
public sealed class PlanPremiumService(IDbSessions sessions, IPlanPremiumRepository repository, TimeProvider clock) : IPlanPremiumService
{
    /// <summary>Desktop texts.</summary>
    public const string InsFirst = "Please Enter Insured Code", HltFirst = "Please Enter Health Code", PlnFirst = "Please Enter Plan Code",
        BadDate = "Please validate your Premium Effective date", Duplicated = "Record Duplicated! Please Change The Corresponding Field",
        NotFound = "Record not found", Saved = "Record Saved...", Deleted = "Record Deleted... ", SearchFirst = "Please search the record",
        InUse = "This record cannot be deleted because it is used in Membership table!";

    /// <inheritdoc />
    public async Task<PremiumSetup> SetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.SetupAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(PremiumDto f)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, Opt(Before(f.InsCode, 4)), Opt(Before(f.PayCode, 4)), Opt(Before(f.HltCode, 4)),
            Opt(Before(f.AgeCode, 4)), Opt(Before(f.PlnCode, 2)));
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<List<PremiumBand>> BandsAsync(string? ins, string? pln)
    {
        await using var s = await sessions.OpenAsync();
        return E(ins) || E(pln) ? [] : await repository.BandsAsync(s, ins!.Trim(), pln!.Trim());
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(PremiumDto p, string userCode)
    {
        if (E(p.InsCode)) throw new BusinessRuleException(InsFirst);
        if (E(p.HltCode)) throw new BusinessRuleException(HltFirst);
        if (E(p.PlnCode)) throw new BusinessRuleException(PlnFirst);
        var eff = ParseDate(p.EffDate) ?? throw new BusinessRuleException(BadDate);
        var rows = Rows(p, eff);
        await using var s = await sessions.BeginAsync();
        var payFilter = Head(p.HltCode, 2) == "S" ? null : Before(p.PayCode, 2);
        if (await repository.ExistsAsync(s, rows[0] with { AgeCode = Head(p.AgeCode, 2) }, payFilter)) throw new ConflictException(Duplicated);
        var today = clock.GetLocalNow().Date;
        foreach (var row in rows) await repository.InsertAsync(s, row, userCode, today);
        await s.CommitAsync();
        return Saved;
    }

    /// <inheritdoc />
    public async Task<string> DeleteAsync(int prmCode, PremiumDto k)
    {
        if (E(k.InsCode)) throw new BusinessRuleException(InsFirst);
        if (prmCode <= 0) throw new BusinessRuleException(SearchFirst);
        await using var s = await sessions.BeginAsync();
        var pay = Head(k.HltCode, 2) == "S" ? null : Before(k.PayCode, 4);
        if (await repository.InUseAsync(s, k.InsCode!.Trim(), (k.HltCode ?? "").Trim(), (k.PlnCode ?? "").Trim(), (k.AgeCode ?? "").Trim(), pay))
            throw new ConflictException(InUse);
        if (await repository.DeleteAsync(s, prmCode) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return Deleted;
    }

    /// <summary>
    /// SavePRM's seven rows. Every row gets the combos' codes (ChkStr(Mid(…, 1, 2)), plan before "-"); a band with a premium gets age code
    /// "0n" and the amount; for status B / C also the supplementary amount of that band (0 when empty, see the screen doc); bands left
    /// empty keep the Age combo's code and no amount.
    /// </summary>
    public static List<PremiumRow> Rows(PremiumDto p, DateTime eff)
    {
        var status = Head(p.SuppStatus, 1) is { Length: 1 } st ? st : "A";
        var version = int.TryParse(p.Version?.Trim(), out var v) ? v : (int?)null;
        var rows = new List<PremiumRow>();
        for (var i = 0; i < 7; i++)
        {
            var amount = Amount(p.Amounts?.ElementAtOrDefault(i));
            var supp = status != "A" && amount is not null ? Amount(p.SuppAmounts?.ElementAtOrDefault(i)) ?? 0m : (decimal?)null;
            rows.Add(new(Head(p.InsCode, 2), Head(p.PayCode, 2), Head(p.HltCode, 2), amount is null ? Head(p.AgeCode, 2) : $"0{i + 1}",
                Before(p.PlnCode, 4), amount, supp, status, eff, version));
        }
        return rows;
    }

    /// <summary>A premium box: empty or a number (thousands separators allowed, as the list shows them).</summary>
    public static bool IsAmountOrEmpty(string? v) => string.IsNullOrWhiteSpace(v) || Amount(v) is not null;

    /// <summary>The masked dd/mm/yyyy box (IsDate); ISO accepted too.</summary>
    public static DateTime? ParseDate(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static decimal? Amount(string? v) =>
        decimal.TryParse(v?.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : null;

    private static bool E(string? v) => string.IsNullOrWhiteSpace(v);

    private static string? Opt(string v) => v.Length == 0 ? null : v;

    /// <summary>ChkStr(Mid(text, 1, n)).</summary>
    private static string Head(string? text, int n) => (text ?? "").Length > n ? text![..n].Trim().ToUpperInvariant() : (text ?? "").Trim().ToUpperInvariant();

    /// <summary>Trim(Mid(text, 1, IIf(InStr(text, "-") = 0, n, InStr(text, "-") - 1))).</summary>
    private static string Before(string? text, int n)
    {
        if (E(text)) return "";
        var at = text!.IndexOf('-');
        return (at < 0 ? text[..Math.Min(n, text.Length)] : text[..at]).Trim();
    }
}
