using Cis.Application.Abstractions;
using Cis.Application.Files;
using Cis.Application.Invoice;
using Cis.Domain.Errors;

namespace Cis.Application.Audit;

/// <summary>Audit review of a case (frmAuditInvoice).</summary>
/// <param name="Case">SP_CISAuditdetails row.</param>
/// <param name="Diagnoses">Diagnosis code → description (TreatmentReason; diagnosis 5 "OTHERS" shows CASDiagnosis5).</param>
public record AuditCase(Dictionary<string, object?> Case, List<InvoiceDiagnosis> Diagnoses);

/// <summary>Data access of the audit review.</summary>
public interface IAuditReviewRepository
{
    /// <summary>cis.Audit_Case.</summary>
    Task<Dictionary<string, object?>?> CaseAsync(IDbSession session, string casNumber);

    /// <summary>cis.InvAdj_Audit.</summary>
    Task UpdateAsync(IDbSession session, string casNumber, AuditReviewRequest request, DateTime auditDate);
}

/// <summary>LG &amp; Audit > Audit review (frmAuditInvoice).</summary>
public interface IAuditReviewService
{
    /// <summary>CmdGo_Click.</summary>
    Task<AuditCase> LoadAsync(string? casNumber);

    /// <summary>CmdUpdate_Click.</summary>
    Task<string> UpdateAsync(string? casNumber, AuditReviewRequest request);

    /// <summary>CmdPrev_Click / CmdNext_Click.</summary>
    Task<string> StepAsync(string? casNumber, bool next);

    /// <summary>The Documents tab: the case's guarantee letter folder (null = Outpatient folder not configured).</summary>
    List<string>? Documents(string? casNumber);

    /// <summary>A document.</summary>
    Task<(byte[] Content, string FileName)> DocumentAsync(string? casNumber, string? fileName);

    /// <summary>CmdCopy_Click: copies a file into the case's guarantee letter folder.</summary>
    Task<string> UploadAsync(string? casNumber, string? fileName, Stream content);
}

/// <inheritdoc />
public sealed class AuditReviewService(IDbSessions sessions, IAuditReviewRepository repository, IClinicInvoiceRepository clinics,
    IInvoiceEnquiryRepository enquiry, IFileDrop files) : IAuditReviewService
{
    /// <summary>Desktop texts.</summary>
    public const string NotFound = "Not Record Found ! ", CaseFirst = "Please Enter Case No", Copied = "File Transfered Successful...";

    /// <summary>addTheFiles' file types.</summary>
    private static readonly string[] Types = [".jpg", ".tif", ".gif", ".doc", ".pdf", ".png", ".msg", ".xlsx", ".dot", ".docx", ".txt"];

    /// <inheritdoc />
    public async Task<AuditCase> LoadAsync(string? casNumber)
    {
        var cas = Case(casNumber);
        await using var s = await sessions.OpenAsync();
        var row = await repository.CaseAsync(s, cas) ?? throw new NotFoundException(NotFound);
        var diagnoses = new List<InvoiceDiagnosis>();
        for (var i = 1; i <= 5; i++)
        {
            var code = InvoiceAdjustmentService.T(row, $"CASTreatReason{i}");
            if (code.Length == 0) continue;
            var d = await clinics.DiagnosisAsync(s, code);
            var desc = d?.Description ?? "";
            // Diagnosis 5 "OTHERS" shows the typed others diagnosis (blank → "OTHERS").
            if (i == 5 && desc.Trim() == "OTHERS" && InvoiceAdjustmentService.T(row, "CASDiagnosis5") is { Length: > 0 } other) desc = other;
            diagnoses.Add(new InvoiceDiagnosis(code, desc, d?.OtherIndicator));
        }
        return new AuditCase(row, diagnoses);
    }

    /// <inheritdoc />
    public async Task<string> UpdateAsync(string? casNumber, AuditReviewRequest r)
    {
        if (string.IsNullOrWhiteSpace(casNumber)) throw new NotFoundException(NotFound);
        var cas = Case(casNumber);
        await using var s = await sessions.OpenAsync();
        if (await repository.CaseAsync(s, cas) is null) throw new NotFoundException(NotFound);
        if (string.IsNullOrWhiteSpace(r.Query)) throw new BusinessRuleException("Please Enter Query! ");
        if (string.IsNullOrWhiteSpace(r.AuditBy)) throw new BusinessRuleException("Please Enter Audit By! ");
        if (string.IsNullOrWhiteSpace(r.AuditDate)) throw new BusinessRuleException("Please Enter Audit Date! ");
        var date = ClinicInvoiceRules.Date(r.AuditDate) ?? throw new BusinessRuleException("Invalid date format!");
        await repository.UpdateAsync(s, cas, r, date);
        return "Audit Details updated ! ";
    }

    /// <inheritdoc />
    public async Task<string> StepAsync(string? casNumber, bool next)
    {
        await using var s = await sessions.OpenAsync();
        return await enquiry.StepAsync(s, (casNumber ?? "").Trim().ToUpperInvariant(), next)
            ?? throw new NotFoundException(next ? InvoiceEnquiryService.Last : InvoiceEnquiryService.First);
    }

    /// <inheritdoc />
    public List<string>? Documents(string? casNumber)
    {
        if (Folder(Case(casNumber)) is not { } dir) return null;
        return Directory.Exists(dir)
            ? Directory.EnumerateFiles(dir).Select(Path.GetFileName).OfType<string>()
                .Where(f => Types.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase)).Order(StringComparer.OrdinalIgnoreCase).ToList()
            : [];
    }

    /// <inheritdoc />
    public async Task<(byte[] Content, string FileName)> DocumentAsync(string? casNumber, string? fileName)
    {
        var cas = Case(casNumber);
        var list = Documents(cas) ?? throw new NotConfiguredException("The outpatient documents folder is not configured (Files:Folders:Outpatient).");
        var name = list.FirstOrDefault(f => string.Equals(f, fileName?.Trim(), StringComparison.OrdinalIgnoreCase)) ?? throw new NotFoundException("File not found.");
        return (await File.ReadAllBytesAsync(Path.Combine(Folder(cas)!, name)), name);
    }

    /// <inheritdoc />
    public async Task<string> UploadAsync(string? casNumber, string? fileName, Stream content)
    {
        if (string.IsNullOrWhiteSpace(casNumber)) throw new BusinessRuleException(CaseFirst);
        var dir = Folder(Case(casNumber)) ?? throw new NotConfiguredException("The outpatient documents folder is not configured (Files:Folders:Outpatient).");
        var name = Path.GetFileName((fileName ?? "").Trim());
        if (name.Length == 0 || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) throw new BusinessRuleException("Please choose a file.");
        Directory.CreateDirectory(dir);
        await using (var target = File.Create(Path.Combine(dir, name))) await content.CopyToAsync(target);
        return Copied;
    }

    /// <summary>DocPath\Outpatient\Guarantee Letter\&lt;case&gt;.</summary>
    private string? Folder(string cas) =>
        files.FolderOf(FileFolders.Outpatient) is { Length: > 0 } root ? Path.Combine(root, "Guarantee Letter", cas) : null;

    /// <summary>A case number (letters and digits only, so it cannot leave the folders).</summary>
    private static string Case(string? casNumber)
    {
        var cas = (casNumber ?? "").Trim().ToUpperInvariant();
        return cas.Length is > 0 and <= 20 && cas.All(char.IsLetterOrDigit) ? cas : throw new NotFoundException(NotFound);
    }
}
