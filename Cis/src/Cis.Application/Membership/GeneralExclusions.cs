using Cis.Application.Abstractions;
using Cis.Application.Files;
using Cis.Domain.Errors;

namespace Cis.Application.Membership;

/// <summary>The group company's general exclusions (frmGeneralBenefitExcEnq).</summary>
/// <param name="GroupCompany">Group company.</param>
/// <param name="Flags">The 24 check boxes by column name (Y = ticked); all false when the company has no row.</param>
/// <param name="Remarks">Remarks.</param>
public record GeneralExclusionsView(string GroupCompany, Dictionary<string, bool> Flags, string Remarks);

/// <summary>Data access (<c>cis.GeneralExclusions_*</c>).</summary>
public interface IGeneralExclusionsRepository
{
    /// <summary>Group company names.</summary>
    Task<List<string>> GroupCompaniesAsync(IDbSession session);

    /// <summary>The company's row, or null.</summary>
    Task<Dictionary<string, object?>?> GetAsync(IDbSession session, string groupCompany);
}

/// <summary>Invoice > General Exclusions (Enquiry).</summary>
public interface IGeneralExclusionsService
{
    /// <summary>ComboList.</summary>
    Task<List<string>> GroupCompaniesAsync();

    /// <summary>cmdSearch_Click → CallExclusions.</summary>
    Task<GeneralExclusionsView> SearchAsync(string? groupCompany);

    /// <summary>Policy Wording tab: files in CLNPolicyWording\&lt;group company&gt; (null when the folder setting is missing).</summary>
    List<string>? PolicyFiles(string? groupCompany);

    /// <summary>lstPolicyWording_Click: one listed file.</summary>
    Task<(byte[] Content, string FileName)> PolicyFileAsync(string? groupCompany, string? fileName);
}

/// <inheritdoc />
public sealed class GeneralExclusionsService(IDbSessions sessions, IGeneralExclusionsRepository repository, IFileDrop files) : IGeneralExclusionsService
{
    /// <summary>The check boxes in form order (column names of GeneralExclusions).</summary>
    public static readonly string[] Columns =
    [
        "Dermatitis", "Gynaecological", "Dental", "Optical", "MedCheckup", "PreEmpCheckup", "Hypertension", "Diabetes", "IHDCAD", "Thyroid",
        "Immunisation", "HRT", "PregnacyTest", "Antenatal", "Abortion", "Infertility", "Contraception", "Genitourinary", "Vitamins", "BloodTest",
        "XRAY", "PapSmear", "Typhoid", "LTMedication",
    ];

    /// <summary>addTheFiles: the extensions listed (any case).</summary>
    public static readonly string[] Extensions = [".jpg", ".tif", ".gif", ".pdf", ".txt", ".xlsx", ".docx"];

    /// <summary>Desktop text.</summary>
    public const string GroupFirst = "Please Select Group Company!";

    /// <inheritdoc />
    public async Task<List<string>> GroupCompaniesAsync()
    {
        await using var s = await sessions.OpenAsync();
        return await repository.GroupCompaniesAsync(s);
    }

    /// <inheritdoc />
    public async Task<GeneralExclusionsView> SearchAsync(string? groupCompany)
    {
        var grp = Group(groupCompany);
        await using var s = await sessions.OpenAsync();
        var row = await repository.GetAsync(s, grp);
        var flags = Columns.ToDictionary(c => c, c => row is not null && string.Equals((row.GetValueOrDefault(c) as string)?.Trim(), "Y", StringComparison.Ordinal));
        return new GeneralExclusionsView(grp, flags, (row?.GetValueOrDefault("Remarks") as string)?.Trim() ?? "");
    }

    /// <inheritdoc />
    public List<string>? PolicyFiles(string? groupCompany)
    {
        var folder = Folder(Group(groupCompany));
        if (folder is null) return null;
        return Directory.Exists(folder)
            ? Directory.EnumerateFiles(folder).Select(Path.GetFileName).OfType<string>()
                .Where(f => Extensions.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase)).Order(StringComparer.OrdinalIgnoreCase).ToList()
            : [];
    }

    /// <inheritdoc />
    public async Task<(byte[] Content, string FileName)> PolicyFileAsync(string? groupCompany, string? fileName)
    {
        var list = PolicyFiles(groupCompany) ?? throw new NotConfiguredException("The policy wording folder is not configured (Files:Folders:PolicyWording).");
        var name = list.FirstOrDefault(f => string.Equals(f, fileName?.Trim(), StringComparison.OrdinalIgnoreCase)) ?? throw new NotFoundException("File not found.");
        return (await File.ReadAllBytesAsync(Path.Combine(Folder(Group(groupCompany))!, name)), name);
    }

    /// <summary>SRVPath\CLNPolicyWording\&lt;group company&gt;; a company name that would leave the folder is refused.</summary>
    private string? Folder(string grp)
    {
        if (files.FolderOf(FileFolders.PolicyWording) is not { Length: > 0 } root) return null;
        var full = Path.GetFullPath(Path.Combine(root, grp));
        return full.StartsWith(Path.GetFullPath(root), StringComparison.OrdinalIgnoreCase) && grp.IndexOfAny(['/', '\\']) < 0 && !grp.Contains("..", StringComparison.Ordinal)
            ? full
            : throw new BusinessRuleException(GroupFirst);
    }

    private static string Group(string? groupCompany) =>
        string.IsNullOrWhiteSpace(groupCompany) ? throw new BusinessRuleException(GroupFirst) : groupCompany.Trim();
}
