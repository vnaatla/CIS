using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.Entity;

/// <summary>A row of lstGrpClinicList / lstDoctorList (code, name).</summary>
/// <param name="Code">GRPClnCode / DocCode.</param>
/// <param name="Name">GRPClnName / DocName.</param>
public record EntityListRow(string Code, string? Name);

/// <summary>Clinic Group fields (frmMaintenanceEntity tab 0).</summary>
/// <param name="Code">GRPClnCode (generated on Add; the record to update on Edit).</param>
/// <param name="Name">Cln Grp. Name.</param>
/// <param name="Initial">Initial Group Name (its first letter starts the generated code).</param>
/// <param name="Add1">Address line 1.</param>
/// <param name="Add2">Address line 2.</param>
/// <param name="Add3">Address line 3.</param>
/// <param name="Contact">Contact Person.</param>
/// <param name="AdmContact">Adm. Contact Person.</param>
/// <param name="Ceo">CEO Name.</param>
/// <param name="Phone">Tel No.</param>
/// <param name="Fax">Fax No.</param>
/// <param name="Email">Email.</param>
/// <param name="Homepage">Homepage.</param>
/// <param name="EffDate">Eff. Date dd/mm/yyyy.</param>
/// <param name="Tax">R&amp;B Tax (Y / N).</param>
/// <param name="Conversion">Conversion (Y / N).</param>
/// <param name="PrePost">Pre&amp;Post (Y / N).</param>
public record ClinicGroupDto(string? Code, string? Name, string? Initial, string? Add1, string? Add2, string? Add3, string? Contact, string? AdmContact,
    string? Ceo, string? Phone, string? Fax, string? Email, string? Homepage, string? EffDate, string? Tax, string? Conversion, string? PrePost);

/// <summary>Save of the Clinic Group tab.</summary>
/// <param name="Group">The fields.</param>
public record ClinicGroupSaveRequest(ClinicGroupDto Group);

/// <summary>Save of the Doctor tab.</summary>
/// <param name="Code">DocCode (Edit only).</param>
/// <param name="Name">Doctor Full Name.</param>
public record DoctorSaveRequest(string? Code, string? Name);

/// <summary>Column lengths of GRPClinic (the desktop's MaxLength); its own checks are in <see cref="EntityGroupDoctorService"/>.</summary>
public sealed class ClinicGroupSaveRequestValidator : RequestValidator<ClinicGroupSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public ClinicGroupSaveRequestValidator()
    {
        RuleFor(r => r.Group).NotNull().ChildRules(g =>
        {
            g.RuleFor(x => x.Initial).MaximumLength(20);
            g.RuleFor(x => x.Phone).MaximumLength(15);
            g.RuleFor(x => x.Fax).MaximumLength(15);
            g.RuleFor(x => x.Name).MaximumLength(50);
            g.RuleFor(x => x.Add1).MaximumLength(50);
            g.RuleFor(x => x.Add2).MaximumLength(50);
            g.RuleFor(x => x.Add3).MaximumLength(50);
            g.RuleFor(x => x.Contact).MaximumLength(50);
            g.RuleFor(x => x.AdmContact).MaximumLength(50);
            g.RuleFor(x => x.Ceo).MaximumLength(50);
            g.RuleFor(x => x.Email).MaximumLength(50);
            g.RuleFor(x => x.Homepage).MaximumLength(50);
        });
    }
}

/// <summary>Doctor name length.</summary>
public sealed class DoctorSaveRequestValidator : RequestValidator<DoctorSaveRequest>
{
    /// <summary>Defines the rules.</summary>
    public DoctorSaveRequestValidator() => RuleFor(r => r.Name).MaximumLength(50);
}

/// <summary>Data access of the Clinic Group and Doctor tabs.</summary>
public interface IEntityGroupDoctorRepository
{
    /// <summary>GRPClinic rows (LIKE code / name).</summary>
    Task<List<EntityListRow>> SearchGroupsAsync(IDbSession session, string? code, string? name);

    /// <summary>One GRPClinic row.</summary>
    Task<ClinicGroupDto?> GetGroupAsync(IDbSession session, string code);

    /// <summary>GenerateGRPClnSeq.</summary>
    Task<string> NextGroupCodeAsync(IDbSession session, string letter);

    /// <summary>Insert or update; returns rows written.</summary>
    Task<int> SaveGroupAsync(IDbSession session, bool insert, ClinicGroupDto group, string userCode, DateTime today);

    /// <summary>True when a HISDB clinic uses the group.</summary>
    Task<bool> GroupInUseAsync(IDbSession session, string code);

    /// <summary>Deletes; returns rows removed.</summary>
    Task<int> DeleteGroupAsync(IDbSession session, string code);

    /// <summary>Doctor rows (LIKE code / name).</summary>
    Task<List<EntityListRow>> SearchDoctorsAsync(IDbSession session, string? code, string? name);

    /// <summary>True when the name is used.</summary>
    Task<bool> DoctorNameExistsAsync(IDbSession session, string name);

    /// <summary>GenerateDocSeq.</summary>
    Task<string> NextDoctorCodeAsync(IDbSession session, string letter);

    /// <summary>Insert or update; returns rows written.</summary>
    Task<int> SaveDoctorAsync(IDbSession session, bool insert, string code, string name, string userCode, DateTime today);

    /// <summary>Deletes; returns rows removed.</summary>
    Task<int> DeleteDoctorAsync(IDbSession session, string code);
}

/// <summary>Maintenance > Entity, Clinic Group and Doctor tabs, with the desktop's checks and messages.</summary>
public interface IEntityGroupDoctorService
{
    /// <summary>SearchClinicGrp (no filter = all rows); "Record not found" when empty.</summary>
    Task<List<EntityListRow>> SearchGroupsAsync(string? code, string? name);

    /// <summary>lstGrpClinicList_Click.</summary>
    Task<ClinicGroupDto> GetGroupAsync(string code);

    /// <summary>SaveClinicGrp → "Record Saved..." / "Record Updated...".</summary>
    Task<string> SaveGroupAsync(ClinicGroupSaveRequest request, bool editing, string userCode);

    /// <summary>DeleteClinicGrp → "Record Deleted...".</summary>
    Task<string> DeleteGroupAsync(string? code);

    /// <summary>SearchDoctor; "Record not found" when empty.</summary>
    Task<List<EntityListRow>> SearchDoctorsAsync(string? code, string? name);

    /// <summary>SaveDoctor.</summary>
    Task<string> SaveDoctorAsync(DoctorSaveRequest request, bool editing, string userCode);

    /// <summary>DeleteDoctor.</summary>
    Task<string> DeleteDoctorAsync(string? code);
}

/// <inheritdoc />
public sealed class EntityGroupDoctorService(IDbSessions sessions, IEntityGroupDoctorRepository repository, TimeProvider clock) : IEntityGroupDoctorService
{
    /// <summary>Desktop texts.</summary>
    public const string GroupName = "Please Enter Clinic Group Name", GroupAddress = "Please Enter Clinic Group Address",
        GroupEffDate = "Please Enter Effective Date", GroupCode = "Please Enter Clinic Group Code",
        GroupInUse = "This record cannot be deleted because it is used in Clinic Table!", DoctorName = "Please Enter Doctor Full Name",
        DoctorCode = "Please Enter Doctor Code", Duplicated = "Record Duplicated! Please Change The Corresponding Field", NotFound = "Record not found",
        Saved = "Record Saved...", Updated = "Record Updated...", Deleted = "Record Deleted... ", SearchFirst = "Please search the record";

    /// <inheritdoc />
    public async Task<List<EntityListRow>> SearchGroupsAsync(string? code, string? name)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchGroupsAsync(s, Up(code), Up(name));
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<ClinicGroupDto> GetGroupAsync(string code)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.GetGroupAsync(s, code.Trim()) ?? throw new NotFoundException(NotFound);
    }

    /// <inheritdoc />
    public async Task<string> SaveGroupAsync(ClinicGroupSaveRequest request, bool editing, string userCode)
    {
        var g = request.Group;
        // The Initial Group Name only builds the code on Add; on Edit its box is locked and existing groups may have none (none locally).
        if (!editing && string.IsNullOrEmpty(g.Initial)) throw new BusinessRuleException(GroupName);
        if (string.IsNullOrEmpty(g.Name)) throw new BusinessRuleException(GroupName);
        if (string.IsNullOrEmpty(g.Add1)) throw new BusinessRuleException(GroupAddress);
        if (string.IsNullOrWhiteSpace(g.EffDate) || g.EffDate.Trim() == "__/__/____") throw new BusinessRuleException(GroupEffDate);
        if (editing && string.IsNullOrWhiteSpace(g.Code)) throw new BusinessRuleException(SearchFirst);

        await using var s = await sessions.BeginAsync();
        var code = editing ? g.Code!.Trim() : await repository.NextGroupCodeAsync(s, g.Initial!.Trim()[..1].ToUpperInvariant());
        var row = CleanGroup(g with { Code = code });
        if (await repository.SaveGroupAsync(s, !editing, row, userCode, clock.GetLocalNow().Date) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return editing ? Updated : Saved;
    }

    /// <inheritdoc />
    public async Task<string> DeleteGroupAsync(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) throw new BusinessRuleException(GroupCode);
        await using var s = await sessions.BeginAsync();
        if (await repository.GroupInUseAsync(s, code.Trim())) throw new ConflictException(GroupInUse);
        if (await repository.DeleteGroupAsync(s, code.Trim()) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return Deleted;
    }

    /// <inheritdoc />
    public async Task<List<EntityListRow>> SearchDoctorsAsync(string? code, string? name)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchDoctorsAsync(s, Up(code), Up(name));
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<string> SaveDoctorAsync(DoctorSaveRequest request, bool editing, string userCode)
    {
        if (string.IsNullOrEmpty(request.Name)) throw new BusinessRuleException(DoctorName);
        if (editing && string.IsNullOrWhiteSpace(request.Code)) throw new BusinessRuleException(SearchFirst);
        await using var s = await sessions.BeginAsync();
        if (!editing && await repository.DoctorNameExistsAsync(s, request.Name.Trim())) throw new ConflictException(Duplicated);
        var code = editing ? request.Code!.Trim() : await repository.NextDoctorCodeAsync(s, request.Name.Trim()[..1].ToUpperInvariant());
        if (await repository.SaveDoctorAsync(s, !editing, code, Up(request.Name)!, userCode, clock.GetLocalNow().Date) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return editing ? Updated : Saved;
    }

    /// <inheritdoc />
    public async Task<string> DeleteDoctorAsync(string? code)
    {
        if (string.IsNullOrWhiteSpace(code)) throw new BusinessRuleException(DoctorCode);
        await using var s = await sessions.BeginAsync();
        if (await repository.DeleteDoctorAsync(s, code.Trim()) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return Deleted;
    }

    /// <summary>SaveClinicGrp's column values: ChkStr (UCase(Trim)) everywhere, NULL for empty optional boxes, Homepage "" when empty, the
    /// first character of the Y/N combos, the date only when valid.</summary>
    public static ClinicGroupDto CleanGroup(ClinicGroupDto g) => new(
        g.Code, Up(g.Name), Up(g.Initial), Up(g.Add1), Up(g.Add2), Up(g.Add3), Up(g.Contact), Up(g.AdmContact), Up(g.Ceo), Up(g.Phone), Up(g.Fax),
        Up(g.Email), (g.Homepage ?? "").Trim().ToUpperInvariant(), ParseDate(g.EffDate)?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        First(g.Tax), First(g.Conversion), First(g.PrePost));

    /// <summary>The masked dd/mm/yyyy box; ISO accepted too.</summary>
    public static DateTime? ParseDate(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static string? Up(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim().ToUpperInvariant();

    private static string First(string? v) => string.IsNullOrWhiteSpace(v) ? "" : v.Trim()[..1].ToUpperInvariant();
}
