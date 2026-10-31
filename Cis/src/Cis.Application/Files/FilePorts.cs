namespace Cis.Application.Files;

/// <summary>Names of the server folders the desktop wrote to (SVRSuper.ini key / desktop usage in brackets).</summary>
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
