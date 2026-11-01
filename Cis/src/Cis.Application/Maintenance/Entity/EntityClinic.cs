using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Common;
using Cis.Application.Documents;
using Cis.Application.Lookups;
using Cis.Domain.Errors;
using FluentValidation;

namespace Cis.Application.Maintenance.Entity;

/// <summary>Clinic Info fields (frmMaintenanceEntity tab 2, HISDB CLN). Combos carry their code letter ("Y", "G", "1" …).</summary>
/// <param name="Code">CLNCode (generated on Add; the clinic to update on Edit; a search filter).</param>
/// <param name="Name">Clinic Name.</param>
/// <param name="Initial">Clinic Initial (its first letter starts the generated code).</param>
/// <param name="GroupCode">Clinic group code.</param>
/// <param name="Type">G - General Practisional / S - Specialised.</param>
/// <param name="BnfType">1 - Standard / 2 - Clinic / 3 - Plan.</param>
/// <param name="Add1">Address 1.</param>
/// <param name="Add2">Address 2.</param>
/// <param name="Add3">Address 3.</param>
/// <param name="PostCode">Post Code.</param>
/// <param name="City">City.</param>
/// <param name="State">State.</param>
/// <param name="ReceivedDate">Received Date dd/mm/yyyy.</param>
/// <param name="Panel">Y / N / D (Declined) / C (Confirmed).</param>
/// <param name="XRay">Y / N.</param>
/// <param name="FormemaRef">Formema Ref.</param>
/// <param name="NoOfDoc">No Of Doctor.</param>
/// <param name="Phone">Tel. No.</param>
/// <param name="Fax">Fax No.</param>
/// <param name="Email">Email.</param>
/// <param name="Homepage">Homepage.</param>
/// <param name="Contact">Contact person (CLNRepresentative).</param>
/// <param name="DocName">Doctor name.</param>
/// <param name="BankAcc">Bank Account.</param>
/// <param name="Bank">Bank.</param>
/// <param name="Branch">Branch.</param>
/// <param name="AccPerson">Account Person.</param>
/// <param name="Hrs24">24 Hours Y / N.</param>
/// <param name="Remark">Remarks.</param>
/// <param name="Ffs">FFS Status Y / N.</param>
/// <param name="FfsPerc">FFS percentage (CLNMCOProcfeePerc).</param>
/// <param name="Location">Location.</param>
/// <param name="OperatingHr">Operating hours.</param>
/// <param name="PayName">Payment clinic name.</param>
/// <param name="PayAdd1">Payment address 1.</param>
/// <param name="PayAdd2">Payment address 2.</param>
/// <param name="PayAdd3">Payment address 3.</param>
/// <param name="Online">Online Status Y / N.</param>
/// <param name="GlFormat">M - Medix / H - Health Connect.</param>
/// <param name="RegNo">Registration No.</param>
public record ClinicDto(string? Code, string? Name, string? Initial, string? GroupCode, string? Type, string? BnfType, string? Add1, string? Add2,
    string? Add3, string? PostCode, string? City, string? State, string? ReceivedDate, string? Panel, string? XRay, string? FormemaRef, string? NoOfDoc,
    string? Phone, string? Fax, string? Email, string? Homepage, string? Contact, string? DocName, string? BankAcc, string? Bank, string? Branch,
    string? AccPerson, string? Hrs24, string? Remark, string? Ffs, string? FfsPerc, string? Location, string? OperatingHr, string? PayName,
    string? PayAdd1, string? PayAdd2, string? PayAdd3, string? Online, string? GlFormat, string? RegNo);

/// <summary>Save / search / export of the Clinic Info tab.</summary>
/// <param name="Clinic">The fields (filters for search and export).</param>
public record ClinicRequest(ClinicDto Clinic);

/// <summary>The tab's combos.</summary>
/// <param name="Groups">HISDB ClinicGrp.</param>
/// <param name="Cities">HISMaintenance CTY.</param>
/// <param name="States">HISDB STA.</param>
public record ClinicSetup(List<LookupOption> Groups, List<LookupOption> Cities, List<LookupOption> States);

/// <summary>Column lengths of CLN (the desktop's MaxLength, cut to the column where it was longer).</summary>
public sealed class ClinicRequestValidator : RequestValidator<ClinicRequest>
{
    /// <summary>Defines the rules.</summary>
    public ClinicRequestValidator() => RuleFor(r => r.Clinic).NotNull().ChildRules(c =>
    {
        c.RuleFor(x => x.Name).MaximumLength(50);
        c.RuleFor(x => x.Initial).MaximumLength(50);
        c.RuleFor(x => x.Add1).MaximumLength(200);
        c.RuleFor(x => x.Add2).MaximumLength(200);
        c.RuleFor(x => x.Add3).MaximumLength(200);
        c.RuleFor(x => x.PostCode).MaximumLength(5);
        c.RuleFor(x => x.FormemaRef).MaximumLength(15);
        c.RuleFor(x => x.NoOfDoc).MaximumLength(3).Matches("^[0-9]*$").WithMessage("No Of Doctor must be a number.");
        c.RuleFor(x => x.Phone).MaximumLength(12);
        c.RuleFor(x => x.Fax).MaximumLength(12);
        c.RuleFor(x => x.Email).MaximumLength(1000);
        c.RuleFor(x => x.Homepage).MaximumLength(50);
        c.RuleFor(x => x.Contact).MaximumLength(50);
        c.RuleFor(x => x.DocName).MaximumLength(50);
        c.RuleFor(x => x.BankAcc).MaximumLength(25);
        c.RuleFor(x => x.Bank).MaximumLength(50);
        c.RuleFor(x => x.Branch).MaximumLength(50);
        c.RuleFor(x => x.AccPerson).MaximumLength(50);
        c.RuleFor(x => x.Remark).MaximumLength(1500);
        c.RuleFor(x => x.FfsPerc).Must(v => string.IsNullOrWhiteSpace(v) || decimal.TryParse(v, NumberStyles.Number, CultureInfo.InvariantCulture, out _))
            .WithMessage("FFS Percentage must be a number.");
        c.RuleFor(x => x.Location).MaximumLength(100);
        c.RuleFor(x => x.OperatingHr).MaximumLength(100);
        c.RuleFor(x => x.PayName).MaximumLength(200);
        c.RuleFor(x => x.PayAdd1).MaximumLength(50);
        c.RuleFor(x => x.PayAdd2).MaximumLength(50);
        c.RuleFor(x => x.PayAdd3).MaximumLength(50);
        c.RuleFor(x => x.RegNo).MaximumLength(50);
    });
}

/// <summary>Data access of the Clinic Info tab.</summary>
public interface IEntityClinicRepository
{
    /// <summary>Combos.</summary>
    Task<ClinicSetup> SetupAsync(IDbSession session);

    /// <summary>CLN rows of the reduced filters.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, ClinicDto filters);

    /// <summary>True when Address 1 is used.</summary>
    Task<bool> AddressExistsAsync(IDbSession session, string add1);

    /// <summary>GenerateCLNSeq.</summary>
    Task<string> NextCodeAsync(IDbSession session, string letter);

    /// <summary>Insert or update the cleaned row; returns rows written.</summary>
    Task<int> SaveAsync(IDbSession session, bool insert, ClinicDto clinic, string userCode, DateTime today);

    /// <summary>True when CISCAS has cases of the clinic.</summary>
    Task<bool> HasCasesAsync(IDbSession session, string code);

    /// <summary>Deletes; returns rows removed.</summary>
    Task<int> DeleteAsync(IDbSession session, string code);
}

/// <summary>Maintenance > Entity, Clinic Info tab.</summary>
public interface IEntityClinicService
{
    /// <summary>CLNComboListing.</summary>
    Task<ClinicSetup> SetupAsync();

    /// <summary>SearchCLN; "Record not found" when empty.</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(ClinicDto filters);

    /// <summary>SaveClinic; returns the message and the code saved.</summary>
    Task<(string Message, string Code)> SaveAsync(ClinicDto clinic, bool editing, string userCode);

    /// <summary>DeleteClinic.</summary>
    Task<string> DeleteAsync(string? code, string? name);

    /// <summary>cmdPrint_Click on Clinic Info → ExportToExcelClinic of the list.</summary>
    Task<FileOutput> ExportAsync(ClinicDto filters);
}

/// <inheritdoc />
public sealed class EntityClinicService(IDbSessions sessions, IEntityClinicRepository repository, IReportExporter exporter, TimeProvider clock)
    : IEntityClinicService
{
    /// <summary>The desktop's checks of SaveClinic, in order (texts kept, leading spaces included).</summary>
    public static readonly IReadOnlyList<(Func<ClinicDto, bool> Missing, string Message)> SaveChecks =
    [
        (c => E(c.Name), "Please Enter Clinic Name"),
        (c => E(c.Initial), "Please Enter Clinic Initial"),
        (c => E(c.Type), "Please Enter Clinic Type"),
        (c => E(c.Add1), "Please Enter Clinic Address"),
        (c => E(c.BnfType), " Please Enter Benefit Type"),
        (c => ParseDate(c.ReceivedDate) is null, "Please Enter Received Date!"),
        (c => E(c.State), " Please Enter State"),
        (c => E(c.PostCode), " Please Enter PostCode"),
        (c => E(c.City), " Please Enter City"),
        (c => E(c.Panel), " Please Enter Panel"),
        (c => E(c.Hrs24), " Please Enter Clinic Operating Hours"),
        (c => E(c.Ffs), " Please Enter FFS Status"),
        (c => c.Ffs?.Trim() == "Y" && (E(c.FfsPerc) || c.FfsPerc!.Trim() == "0"), " Please Enter FFS Percentage"),
        (c => E(c.Location), " Please Enter Location"),
        (c => E(c.OperatingHr), " Please Operating hours"),
    ];

    /// <summary>Desktop texts.</summary>
    public const string Duplicate = "Duplicate Record!", NotFound = "Record not found", Saved = "Record Saved...", Updated = "Record Updated...",
        Deleted = "Record Deleted... ", HasCases = "There are case(s) register under this Clinic !", NameFirst = "Please Enter Clinic Name",
        SearchFirst = "Please search the record", NothingToPrint = "Report cannot be printed without any record.";

    /// <summary>lstClinicInfoList columns (Excel export header).</summary>
    public static readonly IReadOnlyList<ReportColumn> ListColumns =
    [
        new("CLNCode", "Clinic Code"), new("CLNName", "Clinic Name"), new("ClinicGrpCode", "Clinic Group Code"), new("CLNInitial", "Clinic Initial"),
        new("CLNType", "Clinic Type"), new("CLNBNFType", "Benefit Type"), new("CLNAdd1", "Address 1"), new("CLNAdd2", "Address 2"),
        new("CLNAdd3", "Address 3"), new("CLNPostCode", "Post Code"), new("CLNCity", "City"), new("CLNState", "State"),
        new("CLNReceivedDate", "Received Date", ReportFormat.Date), new("CLNPanel", "Panel"), new("CLNXRay", "X-Ray"), new("CLNForRef", "Formema Ref"),
        new("CLNNoofDoc", "No Of Doctor", ReportFormat.Number), new("CLNPhone", "Tel. No"), new("CLNFax", "Fax No"), new("CLNEmail", "Email"),
        new("CLNHomepage", "Homepage"), new("CLNBankAcc", "Bank Account"), new("CLNBankName", "Bank"), new("CLNBankBrn", "Branch"),
        new("CLNAccPerson", "Account Person"), new("CLN24Hrs", "24 Hours"), new("CLNRemark", "Remarks"), new("CLNFFSStatus", "FFS Status"),
        new("CLNOperatingHr", "Operating Hour"), new("CLNPName", "Pay. Clinic Name"), new("CLNPAdd1", "Pay. Add1"), new("CLNPAdd2", "Pay. Add2"),
        new("CLNPAdd3", "Pay. Add3"), new("CLNPPostCode", "Pay. Post Code"), new("CLNPCity", "Pay. City"), new("CLNPState", "Pay. State"),
        new("CLNOnlineStatus", "Online Status"), new("CLNRegNo", "Registration No"),
    ];

    /// <inheritdoc />
    public async Task<ClinicSetup> SetupAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.SetupAsync(s);
    }

    /// <inheritdoc />
    public async Task<List<Dictionary<string, object?>>> SearchAsync(ClinicDto filters)
    {
        await using var s = await sessions.OpenAsync();
        var rows = await repository.SearchAsync(s, ReduceFilters(filters));
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<(string Message, string Code)> SaveAsync(ClinicDto clinic, bool editing, string userCode)
    {
        foreach (var (missing, message) in SaveChecks)
            if (missing(clinic)) throw new BusinessRuleException(message);
        if (editing && E(clinic.Code)) throw new BusinessRuleException(SearchFirst);
        await using var s = await sessions.BeginAsync();
        if (!editing && await repository.AddressExistsAsync(s, clinic.Add1!)) throw new ConflictException(Duplicate);
        var code = editing ? clinic.Code!.Trim() : await repository.NextCodeAsync(s, clinic.Initial!.Trim()[..1].ToUpperInvariant());
        if (await repository.SaveAsync(s, !editing, Clean(clinic with { Code = code }), userCode, clock.GetLocalNow().Date) == 0)
            throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return (editing ? Updated : Saved, code);
    }

    /// <inheritdoc />
    public async Task<string> DeleteAsync(string? code, string? name)
    {
        await using var s = await sessions.BeginAsync();
        var cases = !E(code) && await repository.HasCasesAsync(s, code!.Trim().ToUpperInvariant());
        if (E(name)) throw new BusinessRuleException(NameFirst);
        if (E(code)) throw new BusinessRuleException(SearchFirst);
        if (cases) throw new ConflictException(HasCases);
        if (await repository.DeleteAsync(s, code!.Trim().ToUpperInvariant()) == 0) throw new NotFoundException(SearchFirst);
        await s.CommitAsync();
        return Deleted;
    }

    /// <inheritdoc />
    public async Task<FileOutput> ExportAsync(ClinicDto filters)
    {
        List<Dictionary<string, object?>> rows;
        await using (var s = await sessions.OpenAsync()) rows = await repository.SearchAsync(s, ReduceFilters(filters));
        if (rows.Count == 0) throw new BusinessRuleException(NothingToPrint);
        return exporter.Export(new ReportDocument("Clinic Listing", ListColumns, rows) { FileName = "ClinicListing" }, ReportFileType.Excel);
    }

    /// <summary>SearchCLN's values: code / initial / name / state / city / postcode upper-cased (ChkStr), first letter of the Y/N combos,
    /// the group code before "-" (else its first 4 characters) as a prefix.</summary>
    public static ClinicDto ReduceFilters(ClinicDto f) => Empty with
    {
        Code = Up(f.Code), Initial = Up(f.Initial), Name = Up(f.Name), State = Up(f.State), City = Up(f.City), PostCode = Up(f.PostCode),
        Panel = First(f.Panel), GroupCode = Prefix(f.GroupCode), Type = First(f.Type), Hrs24 = First(f.Hrs24), Ffs = First(f.Ffs),
        Online = First(f.Online), GlFormat = First(f.GlFormat),
    };

    /// <summary>SaveClinic's column values: ChkStr (UCase(Trim)) on the text boxes ("" when empty), the group's first 5 characters, the first
    /// letter of the combos, No Of Doctor / FFS % only when given, Online / GL format "" when not chosen, Registration No as typed.</summary>
    public static ClinicDto Clean(ClinicDto c) => new(
        c.Code, U(c.Name), U(c.Initial), (c.GroupCode ?? "").Trim() is var g && g.Length > 5 ? g[..5].Trim() : (c.GroupCode ?? "").Trim(),
        First(c.Type) ?? "", First(c.BnfType), U(c.Add1), U(c.Add2), U(c.Add3), U(c.PostCode), U(c.City), U(c.State),
        ParseDate(c.ReceivedDate)?.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture), First(c.Panel) ?? "", First(c.XRay) ?? "", U(c.FormemaRef),
        E(c.NoOfDoc) ? null : c.NoOfDoc!.Trim(), U(c.Phone), U(c.Fax), U(c.Email), U(c.Homepage), U(c.Contact), U(c.DocName), U(c.BankAcc), U(c.Bank),
        U(c.Branch), U(c.AccPerson), First(c.Hrs24) ?? "", U(c.Remark), First(c.Ffs) ?? "", E(c.FfsPerc) ? null : c.FfsPerc!.Trim(), U(c.Location),
        U(c.OperatingHr), U(c.PayName), U(c.PayAdd1), U(c.PayAdd2), U(c.PayAdd3), First(c.Online) ?? "", First(c.GlFormat) ?? "", c.RegNo ?? "");

    /// <summary>The masked dd/mm/yyyy box (IsDate); ISO accepted too.</summary>
    public static DateTime? ParseDate(string? text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd"], CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static readonly ClinicDto Empty = new(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null,
        null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

    private static bool E(string? v) => string.IsNullOrEmpty(v);

    private static string U(string? v) => (v ?? "").Trim().ToUpperInvariant();

    private static string? Up(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim().ToUpperInvariant();

    private static string? First(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim()[..1].ToUpperInvariant();

    private static string? Prefix(string? v)
    {
        if (string.IsNullOrWhiteSpace(v)) return null;
        var t = v.Trim();
        var at = t.IndexOf('-');
        return (at < 0 ? t[..Math.Min(4, t.Length)] : t[..at]).Trim();
    }
}
