using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Domain.Errors;
using Cis.Domain.Security;
using FluentValidation;

namespace Cis.Application.Maintenance.CodeTables;

/// <summary>Values of one record, by column name (the form's text boxes).</summary>
/// <param name="Values">Column → typed value.</param>
public record CodeTableSaveRequest(Dictionary<string, string?> Values);

/// <summary>Shape check only; the field rules (desktop messages, order) are applied by <see cref="CodeTableService"/> from the catalog.</summary>
public sealed class CodeTableSaveRequestValidator : RequestValidator<CodeTableSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public CodeTableSaveRequestValidator() => RuleFor(r => r.Values).NotNull().WithMessage("Values are required.");
}

/// <summary>A field as the screen needs it.</summary>
/// <param name="Column">Column name (value key).</param>
/// <param name="Label">Caption.</param>
/// <param name="MaxLength">Maximum length.</param>
/// <param name="Required">Mandatory (shown with *).</param>
/// <param name="Kind">Entry kind.</param>
/// <param name="Searchable">Used by Search (shown in bold in the desktop).</param>
/// <param name="IsKey">The record's code (read-only when editing).</param>
/// <param name="HasChoices">Options come from <c>choices/{column}</c>.</param>
public record CodeFieldView(string Column, string Label, int MaxLength, bool Required, FieldKind Kind, bool Searchable, bool IsKey, bool HasChoices);

/// <summary>A code table as the screen needs it.</summary>
/// <param name="Key">URL key.</param>
/// <param name="LegacyForm">VB6 form.</param>
/// <param name="Tab">Tab caption.</param>
/// <param name="ListReportKey">Print key (null = no Print button).</param>
/// <param name="Fields">Fields in order.</param>
public record CodeTableView(string Key, string LegacyForm, string Tab, string? ListReportKey, IReadOnlyList<CodeFieldView> Fields);

/// <summary>Result of a save / delete: the desktop's confirmation text.</summary>
/// <param name="Message">e.g. "Record Saved...".</param>
public record CodeTableResult(string Message);

/// <summary>Data access of the code tables.</summary>
public interface ICodeTableRepository
{
    /// <summary>Rows matching the searchable values (all rows when none), at most <paramref name="max"/>.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, CodeTableDefinition table, IReadOnlyDictionary<string, string> filters, int max);

    /// <summary>True when a row with the code exists.</summary>
    Task<bool> ExistsAsync(IDbSession session, CodeTableDefinition table, string code);

    /// <summary>Inserts a row with the audit columns.</summary>
    Task InsertAsync(IDbSession session, CodeTableDefinition table, IReadOnlyDictionary<string, string?> values, string userCode, DateTime date);

    /// <summary>Updates the first row with the code (the desktop's keyset recordset); returns the rows changed.</summary>
    Task<int> UpdateFirstAsync(IDbSession session, CodeTableDefinition table, string code, IReadOnlyDictionary<string, string?> values, string userCode, DateTime date);

    /// <summary>Deletes every row with the code; returns the rows deleted.</summary>
    Task<int> DeleteAsync(IDbSession session, CodeTableDefinition table, string code);

    /// <summary>Options of a choice field.</summary>
    Task<List<Lookups.LookupOption>> ChoicesAsync(IDbSession session, ChoiceSource source);
}

/// <summary>Search / add / edit / delete of the maintenance code tables with the desktop's messages.</summary>
public interface ICodeTableService
{
    /// <summary>The table's fields for the screen.</summary>
    CodeTableView Describe(string key, string? userAccess);

    /// <summary>The desktop's Search (all rows when no searchable value is given).</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(string key, IReadOnlyDictionary<string, string> filters, string? userAccess);

    /// <summary>Options of a choice field.</summary>
    Task<List<Lookups.LookupOption>> ChoicesAsync(string key, string column, string? userAccess);

    /// <summary>Add + Save.</summary>
    Task<CodeTableResult> CreateAsync(string key, CodeTableSaveRequest request, string userCode, string? userAccess);

    /// <summary>Edit + Save of the record with <paramref name="code"/>.</summary>
    Task<CodeTableResult> UpdateAsync(string key, string code, CodeTableSaveRequest request, string userCode, string? userAccess);

    /// <summary>Delete of the record(s) with <paramref name="code"/>.</summary>
    Task<CodeTableResult> DeleteAsync(string key, string code, string? userAccess);
}

/// <inheritdoc />
public sealed class CodeTableService(IDbSessions sessions, ICodeTableRepository repository, TimeProvider clock) : ICodeTableService
{
    /// <summary>Most rows a search returns.</summary>
    public const int MaxRows = 5000;

    /// <summary>Desktop texts.</summary>
    public const string Saved = "Record Saved...", Updated = "Record Updated...", Deleted = "Record Deleted... ",
        Duplicated = "Record Duplicated! Please Change The Corresponding Field", SearchFirst = "Please search the record";

    /// <inheritdoc />
    public CodeTableView Describe(string key, string? userAccess)
    {
        var table = Table(key, userAccess);
        return new CodeTableView(table.Key, table.LegacyForm, table.Tab, table.ListReportKey,
            table.Fields.Select(f => new CodeFieldView(f.Column, f.Label, f.MaxLength, f.Required is not null, f.Kind, f.Search != SearchMatch.None, f.IsKey, f.Choices is not null)).ToList());
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(string key, IReadOnlyDictionary<string, string> filters, string? userAccess)
    {
        var table = Table(key, userAccess);
        var used = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in table.Fields.Where(f => f.Search != SearchMatch.None))
        {
            var value = filters.FirstOrDefault(kv => string.Equals(kv.Key, field.Column, StringComparison.OrdinalIgnoreCase)).Value?.Trim();
            if (string.IsNullOrEmpty(value)) continue;
            used[field.Column] = field.Search switch
            {
                SearchMatch.EqualsFirst3 => value.Length > 3 ? value[..3].Trim() : value,
                SearchMatch.EqualsFirstChar => value[..1],
                _ => value,
            };
        }
        await using var session = await sessions.OpenAsync();
        return await repository.SearchAsync(session, table, used, MaxRows);
    }

    /// <inheritdoc />
    public async Task<List<Lookups.LookupOption>> ChoicesAsync(string key, string column, string? userAccess)
    {
        var table = Table(key, userAccess);
        var source = table.Fields.FirstOrDefault(f => string.Equals(f.Column, column, StringComparison.OrdinalIgnoreCase))?.Choices
            ?? throw new NotFoundException($"Field '{column}' has no options.");
        await using var session = await sessions.OpenAsync();
        return await repository.ChoicesAsync(session, source);
    }

    /// <inheritdoc />
    public async Task<CodeTableResult> CreateAsync(string key, CodeTableSaveRequest request, string userCode, string? userAccess)
    {
        var table = Table(key, userAccess);
        var values = Validate(table, request.Values, editing: false);
        await using var session = await sessions.BeginAsync();
        if (await repository.ExistsAsync(session, table, values[table.KeyField.Column]!)) throw new ConflictException(Duplicated);
        await repository.InsertAsync(session, table, values, userCode, Today());
        await session.CommitAsync();
        return new CodeTableResult(Saved);
    }

    /// <inheritdoc />
    public async Task<CodeTableResult> UpdateAsync(string key, string code, CodeTableSaveRequest request, string userCode, string? userAccess)
    {
        var table = Table(key, userAccess);
        var input = new Dictionary<string, string?>(request.Values, StringComparer.OrdinalIgnoreCase) { [table.KeyField.Column] = code };
        var values = Validate(table, input, editing: true);
        await using var session = await sessions.BeginAsync();
        if (await repository.UpdateFirstAsync(session, table, code.Trim(), values, userCode, Today()) == 0) throw new NotFoundException(SearchFirst);
        await session.CommitAsync();
        return new CodeTableResult(Updated);
    }

    /// <inheritdoc />
    public async Task<CodeTableResult> DeleteAsync(string key, string code, string? userAccess)
    {
        var table = Table(key, userAccess);
        if (string.IsNullOrWhiteSpace(code)) throw new BusinessRuleException(table.KeyField.Required ?? SearchFirst);
        await using var session = await sessions.BeginAsync();
        if (await repository.DeleteAsync(session, table, code.Trim()) == 0) throw new NotFoundException(SearchFirst);
        await session.CommitAsync();
        return new CodeTableResult(Deleted);
    }

    /// <summary>The desktop's Save checks, field by field in form order; returns the values to store.</summary>
    public static Dictionary<string, string?> Validate(CodeTableDefinition table, IReadOnlyDictionary<string, string?> input, bool editing)
    {
        var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in table.Fields)
        {
            var raw = input.FirstOrDefault(kv => string.Equals(kv.Key, field.Column, StringComparison.OrdinalIgnoreCase)).Value?.Trim() ?? "";
            if (raw.Length == 0)
            {
                if (field.Required is not null) throw new BusinessRuleException(field.Required);
                values[field.Column] = field.Kind == FieldKind.Text ? "" : null;
                continue;
            }
            if (field.MinLength > 0 && raw.Length < field.MinLength && !(editing && field.IsKey)) throw new BusinessRuleException(field.MinMessage!);
            if (raw.Length > field.MaxLength && field.Kind != FieldKind.ChoiceFirstChar)
                throw new BusinessRuleException($"{field.Label} must be at most {field.MaxLength} characters.");
            values[field.Column] = field.Kind switch
            {
                FieldKind.Number or FieldKind.MoneyText when !decimal.TryParse(raw, NumberStyles.Number, CultureInfo.InvariantCulture, out _) =>
                    throw new BusinessRuleException(field.Required ?? $"{field.Label} must be a number."),
                FieldKind.MoneyText => decimal.Parse(raw, NumberStyles.Number, CultureInfo.InvariantCulture).ToString("#,##0.00", CultureInfo.InvariantCulture),
                FieldKind.ChoiceFirstChar => raw[..1],
                _ => raw,
            };
        }
        return values;
    }

    private DateTime Today() => clock.GetLocalNow().Date;

    private static CodeTableDefinition Table(string key, string? userAccess)
    {
        var table = CodeTableCatalog.Find(key) ?? throw new NotFoundException($"Unknown code table '{key}'.");
        return AccessRight.Has(userAccess, table.Access) ? table : throw new ForbiddenException();
    }
}
