using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.Entity;

/// <summary>A lstGrpComCLN row (VIEW_GrpCom_Clinic).</summary>
/// <param name="GrpComCode">Group company code.</param>
/// <param name="GrpCoName">Group company name.</param>
/// <param name="GrpComClnCode">Clinic code.</param>
/// <param name="ClnName">Clinic name.</param>
/// <param name="ClnAddress">Address 1 + 2 + 3.</param>
/// <param name="ClnPostCode">Post code.</param>
/// <param name="ClnState">State.</param>
/// <param name="ClnCity">City.</param>
/// <param name="ClnLocation">Location.</param>
/// <param name="Cln24Hrs">24 hours.</param>
/// <param name="ClnPhone">Phone.</param>
/// <param name="ClnFax">Fax.</param>
public record GroupCompanyClinic(string GrpComCode, string? GrpCoName, string GrpComClnCode, string? ClnName, string? ClnAddress, string? ClnPostCode,
    string? ClnState, string? ClnCity, string? ClnLocation, string? Cln24Hrs, string? ClnPhone, string? ClnFax);

/// <summary>Search answer: the rows shown (at most <see cref="EntityGrpComAccessService.MaxRows"/>) and how many match.</summary>
/// <param name="Rows">Rows.</param>
/// <param name="Total">All matching rows.</param>
public record GroupCompanyClinicPage(List<GroupCompanyClinic> Rows, long Total);

/// <summary>Group company and clinic combos (Check1).</summary>
/// <param name="GroupCompanies">CISDB GRPCompany.</param>
/// <param name="Clinics">HISDB CLN.</param>
public record GroupCompanyClinicLookups(List<LookupOption> GroupCompanies, List<LookupOption> Clinics);

/// <summary>Save of the Group Company Clinic tab (combo texts "CODE - NAME" or codes).</summary>
/// <param name="GroupCompany">Group company.</param>
/// <param name="Clinic">Clinic.</param>
/// <param name="OldGroupCompany">Edit: the selected row's group company.</param>
/// <param name="OldClinic">Edit: the selected row's clinic.</param>
public record GroupCompanyClinicSaveRequest(string? GroupCompany, string? Clinic, string? OldGroupCompany, string? OldClinic);

/// <summary>A CLNAccessControl row (lstClinicAccess).</summary>
/// <param name="ClinicCode">Clinic code.</param>
/// <param name="ClnName">Clinic name.</param>
/// <param name="EffDateFrom">Effective from.</param>
/// <param name="EffDateTo">Effective to.</param>
/// <param name="CreatedBy">User.</param>
/// <param name="CreatedDate">Created.</param>
public record ClinicAccess(string ClinicCode, string? ClnName, DateTime? EffDateFrom, DateTime? EffDateTo, string? CreatedBy, DateTime? CreatedDate);

/// <summary>Save of the Clinic Access tab.</summary>
/// <param name="ClinicCode">Clinic code.</param>
/// <param name="EffDateFrom">Effective Date From dd/mm/yyyy.</param>
/// <param name="EffDateTo">Effective Date To dd/mm/yyyy.</param>
public record ClinicAccessSaveRequest(string? ClinicCode, string? EffDateFrom, string? EffDateTo);

/// <summary>A frmSearchCLN row.</summary>
/// <param name="Code">CLNCode.</param>
/// <param name="Name">CLNName.</param>
/// <param name="Location">CLNCity.</param>
/// <param name="DocName">CLNDocName.</param>
public record ClinicFound(string Code, string? Name, string Location, string DocName);

/// <summary>Lengths of GRPCOMClinic.</summary>
public sealed class GroupCompanyClinicSaveRequestValidator : RequestValidator<GroupCompanyClinicSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public GroupCompanyClinicSaveRequestValidator()
    {
        RuleFor(r => r.GroupCompany).MaximumLength(300);
        RuleFor(r => r.Clinic).MaximumLength(300);
    }
}

/// <summary>Length of CLNAccessControl.ClinicCode.</summary>
public sealed class ClinicAccessSaveRequestValidator : RequestValidator<ClinicAccessSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public ClinicAccessSaveRequestValidator() => RuleFor(r => r.ClinicCode).MaximumLength(10);
}

/// <summary>Data access of the Group Company Clinic and Clinic Access tabs.</summary>
public interface IEntityGrpComAccessRepository
{
    /// <summary>Combos.</summary>
    Task<GroupCompanyClinicLookups> LookupsAsync(IDbSession session);

    /// <summary>Rows (capped) and total.</summary>
    Task<GroupCompanyClinicPage> SearchAsync(IDbSession session, string? grpCom, string? cln, int page, int pageSize, string? sortKey, bool sortDesc);

    /// <summary>True when the pair exists.</summary>
    Task<bool> ExistsAsync(IDbSession session, string grpCom, string cln);

    /// <summary>Insert / update; rows written.</summary>
    Task<int> SaveAsync(IDbSession session, bool insert, string? oldGrpCom, string? oldCln, string grpCom, string cln, string userCode, DateTime today);

    /// <summary>Delete; rows removed.</summary>
    Task<int> DeleteAsync(IDbSession session, string grpCom, string cln);

    /// <summary>Latest 100 access rows.</summary>
    Task<List<ClinicAccess>> AccessListAsync(IDbSession session);

    /// <summary>The clinic's name (null = no such clinic).</summary>
    Task<string?> ClinicNameAsync(IDbSession session, string code);

    /// <summary>Insert an access row.</summary>
    Task InsertAccessAsync(IDbSession session, string code, DateTime from, DateTime to, string userCode, DateTime now);

    /// <summary>frmSearchCLN.</summary>
    Task<List<ClinicFound>> FindClinicsAsync(IDbSession session, string name);
}

/// <summary>Maintenance > Entity, Group Company Clinic and Clinic Access tabs, and frmSearchCLN.</summary>
public interface IEntityGrpComAccessService
{
    /// <summary>GRPCOMCLNComboListing.</summary>
    Task<GroupCompanyClinicLookups> LookupsAsync();

    /// <summary>SearchGrpComCLN; "Record not found" when empty.</summary>
    Task<GroupCompanyClinicPage> SearchAsync(string? groupCompany, string? clinic, int page = 0, int pageSize = 50, string? sortKey = null, bool sortDesc = false);

    /// <summary>SaveGrpComClinic.</summary>
    Task<string> SaveAsync(GroupCompanyClinicSaveRequest request, bool editing, string userCode);

    /// <summary>DeleteGrpComClinic.</summary>
    Task<string> DeleteAsync(string? groupCompany, string? clinic);

    /// <summary>ClinicAccessListing; "Record not found" when empty.</summary>
    Task<List<ClinicAccess>> AccessListAsync();

    /// <summary>SaveClnAccess.</summary>
    Task<string> SaveAccessAsync(ClinicAccessSaveRequest request, string userCode);

    /// <summary>frmSearchCLN.SearchClinic.</summary>
    Task<List<ClinicFound>> FindClinicsAsync(string? name);
}

/// <inheritdoc />
public sealed class EntityGrpComAccessService(IDbSessions sessions, IEntityGrpComAccessRepository repository, TimeProvider clock) : IEntityGrpComAccessService
{
    /// <summary>Rows shown at most (the desktop listed the whole view, ~1.7 million rows locally).</summary>
    public const int MaxRows = 1000;

    /// <summary>Desktop texts (spelling kept).</summary>
    public const string GroupCompanyFirst = "Please Enter Group Company Name", ClinicFirst = "Please Enter Clinic Code",
        Duplicated = "Record Duplicated! Please Change The Corresponding Field", NotFound = "Record not found", Saved = "Record Saved...",
        Updated = "Record Updated...", Deleted = "Record Deleted... ", SearchFirst = "Please search the record",
        FromInvalid = "Invalid Effective Date From", ToInvalid = "InvalidEffective Date To",
        FromAfterTo = "Effective Date From cannot be more than Effective Date To!", NoClinic = "Clinic Code does not exists, please enter again!",
        NameToSearch = "Please Enter Clinic Name to search";

    /// <inheritdoc />
    public async Task<GroupCompanyClinicLookups> LookupsAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.LookupsAsync(s);
    }

    /// <inheritdoc />
    public async Task<GroupCompanyClinicPage> SearchAsync(string? groupCompany, string? clinic, int page = 0, int pageSize = 50, string? sortKey = null, bool sortDesc = false)
    {
        await using var s = await sessions.OpenAsync();
        var result = await repository.SearchAsync(s, CodeOf(groupCompany), CodeOf(clinic), Math.Max(0, page), Math.Clamp(pageSize, 1, MaxRows), sortKey, sortDesc);
        return result.Total == 0 ? throw new NotFoundException(NotFound) : result;
    }

    /// <inheritdoc />
    public async Task<string> SaveAsync(GroupCompanyClinicSaveRequest r, bool editing, string userCode)
    {
        if (string.IsNullOrEmpty(r.GroupCompany)) throw new BusinessRuleException(GroupCompanyFirst);
        if (string.IsNullOrEmpty(r.Clinic)) throw new BusinessRuleException(ClinicFirst);
        if (editing && (CodeOf(r.OldGroupCompany) is null || CodeOf(r.OldClinic) is null)) throw new BusinessRuleException(SearchFirst);
        var (grpCom, cln) = (CodeOf(r.GroupCompany)!, CodeOf(r.Clinic)!);
        await using var s = await sessions.BeginAsync();
        if (await repository.ExistsAsync(s, grpCom, cln)) throw new ConflictException(Duplicated);
        if (await repository.SaveAsync(s, !editing, CodeOf(r.OldGroupCompany), CodeOf(r.OldClinic), grpCom, cln, userCode, clock.GetLocalNow().Date) == 0)
            throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return editing ? Updated : Saved;
    }

    /// <inheritdoc />
    public async Task<string> DeleteAsync(string? groupCompany, string? clinic)
    {
        if (CodeOf(groupCompany) is not { } grpCom || CodeOf(clinic) is not { } cln) throw new BusinessRuleException(SearchFirst);
        await using var s = await sessions.BeginAsync();
        if (await repository.DeleteAsync(s, grpCom, cln) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return Deleted;
    }

    /// <inheritdoc />
    public async Task<List<ClinicAccess>> AccessListAsync()
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.AccessListAsync(s);
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<string> SaveAccessAsync(ClinicAccessSaveRequest r, string userCode)
    {
        if (string.IsNullOrEmpty(r.ClinicCode)) throw new BusinessRuleException(ClinicFirst);
        var from = SmallDate(r.EffDateFrom) ?? throw new BusinessRuleException(FromInvalid);
        var to = SmallDate(r.EffDateTo) ?? throw new BusinessRuleException(ToInvalid);
        if (from > to) throw new BusinessRuleException(FromAfterTo);
        await using var s = await sessions.BeginAsync();
        if (await repository.ClinicNameAsync(s, r.ClinicCode.Trim()) is null) throw new NotFoundException(NoClinic);
        await repository.InsertAccessAsync(s, r.ClinicCode.Trim().ToUpperInvariant(), from, to, userCode, clock.GetLocalNow().DateTime);
        await s.CommitAsync();
        return Saved;
    }

    /// <inheritdoc />
    public async Task<List<ClinicFound>> FindClinicsAsync(string? name)
    {
        if (string.IsNullOrEmpty(name)) throw new BusinessRuleException(NameToSearch);
        await using var s = await sessions.OpenAsync();
        var rows = await repository.FindClinicsAsync(s, name.Trim());
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <summary>The combo's code: text before "-" (else its first 10 characters), trimmed; null when empty.</summary>
    public static string? CodeOf(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        var at = text.IndexOf('-');
        return (at < 0 ? text[..Math.Min(10, text.Length)] : text[..at]).Trim() is { Length: > 0 } code ? code : null;
    }

    /// <summary>A date that fits CLNAccessControl's smalldatetime columns (01/01/1900 – 06/06/2079), else null.</summary>
    public static DateTime? SmallDate(string? text) => ParseDate(text) is { } d && d >= new DateTime(1900, 1, 1) && d <= new DateTime(2079, 6, 6) ? d : null;

    /// <summary>The masked dd/mm/yyyy box (IsDate); ISO accepted too.</summary>
    public static DateTime? ParseDate(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;
}
