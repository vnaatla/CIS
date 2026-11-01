using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.BenefitClinic;

/// <summary>A saved RM Payable Fee of a category (STDClinicBNFSpe row).</summary>
/// <param name="BnfCode">BNFCode (category code).</param>
/// <param name="RmPerVisit">CLNBNFRMPerVisit.</param>
public record ClinicBenefitFee(string BnfCode, string? RmPerVisit);

/// <summary>Save of the screen: card type, effective date and the grid's RM Payable Fee per category code.</summary>
/// <param name="CardType">BNF Card Type (Classic / Silver / Gold).</param>
/// <param name="EffDate">Eff Date as typed (dd/mm/yyyy).</param>
/// <param name="Fees">RM Payable Fee by category code; missing / empty = no fee.</param>
public record ClinicBenefitSaveRequest(string? CardType, string? EffDate, Dictionary<string, string?>? Fees);

/// <summary>Column lengths of STDClinicBNFSpe; the desktop's own checks (and their order) are in <see cref="BenefitClinicService"/>.</summary>
public sealed class ClinicBenefitSaveRequestValidator : RequestValidator<ClinicBenefitSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public ClinicBenefitSaveRequestValidator()
    {
        RuleFor(r => r.CardType).MaximumLength(20).WithMessage("BNF Card Type must be at most 20 characters.");
        RuleForEach(r => r.Fees).Must(f => (f.Value?.Trim().Length ?? 0) <= 50).WithMessage("RM Payable Fee must be at most 50 characters.");
    }
}

/// <summary>Data access of Benefit Clinic.</summary>
public interface IBenefitClinicRepository
{
    /// <summary>BNFCategoryNew (the grid rows).</summary>
    Task<List<LookupOption>> CategoriesAsync(IDbSession session);

    /// <summary>STDClinicBNFSpe rows of the card type + date.</summary>
    Task<List<ClinicBenefitFee>> SearchAsync(IDbSession session, string cardType, DateTime effDate);

    /// <summary>True when the card type + date already has rows.</summary>
    Task<bool> ExistsAsync(IDbSession session, string cardType, DateTime effDate);

    /// <summary>Inserts one STDClinicBNFSpe row.</summary>
    Task InsertAsync(IDbSession session, string cardType, DateTime effDate, string bnfCode, string? rmPerVisit, string userCode, DateTime today);
}

/// <summary>Maintenance > Benefit > Clinic with the desktop's checks and messages.</summary>
public interface IBenefitClinicService
{
    /// <summary>The fixed card types of cboCardType.</summary>
    IReadOnlyList<string> CardTypes { get; }

    /// <summary>Grid rows: every benefit category.</summary>
    Task<List<LookupOption>> CategoriesAsync();

    /// <summary>SearchClinicBNF: needs card type and a valid date; "Record not found!" when nothing is saved.</summary>
    Task<List<ClinicBenefitFee>> SearchAsync(string? cardType, string? effDate);

    /// <summary>SaveClinicBNF: checks, duplicate check, one row per category → "Record Saved...".</summary>
    Task<string> SaveAsync(ClinicBenefitSaveRequest request, string userCode);
}

/// <inheritdoc />
public sealed class BenefitClinicService(IDbSessions sessions, IBenefitClinicRepository repository, TimeProvider clock) : IBenefitClinicService
{
    /// <summary>Desktop texts (spelling kept).</summary>
    public const string SearchType = "Please key in BNF Card Type!", SearchDate = "Please key in Effective Date!",
        SearchValidDate = "Please key in VALID Effective Date!", SaveType = "Please Enter Plan Code!",
        SaveDate = "Please Enter (Valid )effectice Date!", Duplicated = "Record Duplicated! Please Change The Corresponding Field",
        NotFound = "Record not found!", Saved = "Record Saved...";

    /// <inheritdoc />
    public IReadOnlyList<string> CardTypes { get; } = ["Classic", "Silver", "Gold"];

    /// <inheritdoc />
    public async Task<List<LookupOption>> CategoriesAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.CategoriesAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<ClinicBenefitFee>> SearchAsync(string? cardType, string? effDate)
    {
        if (string.IsNullOrWhiteSpace(cardType)) throw new BusinessRuleException(SearchType);
        if (string.IsNullOrWhiteSpace(effDate)) throw new BusinessRuleException(SearchDate);
        var date = ParseDate(effDate) ?? throw new BusinessRuleException(SearchValidDate);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, cardType.Trim(), date);
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(ClinicBenefitSaveRequest request, string userCode)
    {
        if (string.IsNullOrWhiteSpace(request.CardType)) throw new BusinessRuleException(SaveType);
        var date = ParseDate(request.EffDate) ?? throw new BusinessRuleException(SaveDate);
        var cardType = request.CardType.Trim();
        await using var s = await sessions.BeginAsync();
        if (await repository.ExistsAsync(s, cardType, date)) throw new ConflictException(Duplicated);
        var today = clock.GetLocalNow().Date;
        foreach (var category in await repository.CategoriesAsync(s))
            await repository.InsertAsync(s, cardType, date, category.Code, Fee(request.Fees, category.Code), userCode, today);
        await s.CommitAsync();
        return Saved;
    }

    /// <summary>FGrid_Click keeps only numeric entries; ChkStr = UCase(Trim()). Empty → NULL (the desktop left the column unset).</summary>
    public static string? Fee(Dictionary<string, string?>? fees, string code)
    {
        if (fees is null || !fees.TryGetValue(code, out var raw) || string.IsNullOrWhiteSpace(raw)) return null;
        var v = raw.Trim().ToUpperInvariant();
        return decimal.TryParse(v, NumberStyles.Number, CultureInfo.InvariantCulture, out _) ? v : null;
    }

    /// <summary>The masked dd/mm/yyyy box (IsDate); ISO accepted too.</summary>
    public static DateTime? ParseDate(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;
}
