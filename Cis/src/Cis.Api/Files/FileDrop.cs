using Microsoft.Extensions.Options;

namespace Cis.Api.Files;

/// <summary>
/// Server folders the desktop wrote to (configuration section <c>Files</c>, keys = the SVRSuper.ini settings they replace). Empty =
/// files are only downloaded by the browser, not copied to a share.
/// </summary>
public class FileOptions
{
    /// <summary>Configuration section name.</summary>
    public const string SectionName = "Files";

    /// <summary>Folder per kind of file, e.g. <c>ClaimAscii</c> → <c>\\svr-doc\ClinicalClaimASCII</c>.</summary>
    public Dictionary<string, string> Folders { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

/// <summary>Names of the configured folders (SVRSuper.ini key / desktop usage in brackets).</summary>
public static class FileFolders
{
    /// <summary>Claim ASCII exports (SRVPath &amp; "ClinicalClaimASCII").</summary>
    public const string ClaimAscii = "ClaimAscii";
    /// <summary>Bordx invoices per client (BordxPath, e.g. "SEB\Invoices").</summary>
    public const string Bordx = "Bordx";
    /// <summary>MCO invoice files (MCOPath &amp; "CIS").</summary>
    public const string McoInvoice = "McoInvoice";
    /// <summary>Payment vouchers (LocreportPath &amp; "PaymentVoucher").</summary>
    public const string PaymentVoucher = "PaymentVoucher";
    /// <summary>Outpatient documents (DocPath &amp; "Outpatient").</summary>
    public const string Outpatient = "Outpatient";
    /// <summary>Clinic payment vouchers (DocPath &amp; "ClinicPVS").</summary>
    public const string ClinicPvs = "ClinicPvs";
    /// <summary>Clinic policy wording (SRVPath &amp; "CLNPolicyWording").</summary>
    public const string PolicyWording = "PolicyWording";
    /// <summary>LG letters (LocreportPath &amp; "CIS\LG").</summary>
    public const string LgLetters = "LgLetters";
}

/// <summary>Copies generated files to the server folders the desktop used.</summary>
public interface IFileDrop
{
    /// <summary>Saves the file in the folder configured for <paramref name="folder"/>; returns the full path, or null when that folder is not configured.</summary>
    Task<string?> SaveAsync(string folder, string fileName, byte[] content);

    /// <summary>The configured folder, or null.</summary>
    string? FolderOf(string folder);
}

/// <inheritdoc />
public sealed class FileDrop(IOptions<FileOptions> options) : IFileDrop
{
    /// <inheritdoc />
    public string? FolderOf(string folder) =>
        options.Value.Folders.TryGetValue(folder, out var path) && !string.IsNullOrWhiteSpace(path) ? path : null;

    /// <inheritdoc />
    public async Task<string?> SaveAsync(string folder, string fileName, byte[] content)
    {
        if (FolderOf(folder) is not { } root) return null;
        var name = Path.GetFileName(fileName);
        var fullRoot = Path.GetFullPath(root);
        var path = Path.GetFullPath(Path.Combine(fullRoot, name));
        if (name.Trim('.', ' ').Length == 0 || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
            !string.Equals(Path.GetDirectoryName(path), fullRoot.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException($"Invalid file name '{fileName}'.", nameof(fileName));
        Directory.CreateDirectory(fullRoot);
        await File.WriteAllBytesAsync(path, content);
        return path;
    }
}
