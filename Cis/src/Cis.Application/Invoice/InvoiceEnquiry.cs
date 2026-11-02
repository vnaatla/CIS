using System.Globalization;
using System.Text.RegularExpressions;
using Cis.Application.Abstractions;
using Cis.Application.Files;
using Cis.Domain.Errors;

namespace Cis.Application.Invoice;

/// <summary>The stored worksheet of a case (WSInfo): the seven worksheet tables in grid order and the Claim totals.</summary>
/// <param name="Rows">31 benefit rows.</param>
/// <param name="Claim">Claim (CLMTotalActual, CLMFinalApprove, CLMPayableByMedix2C / 2M, CLMMCOProcFee …).</param>
public record StoredWorksheet(List<WorksheetRow> Rows, Dictionary<string, object?>? Claim);

/// <summary>Enquiry of one case (FrmInvoiceEnq).</summary>
/// <param name="Case">SP_INV_CLAIM row.</param>
/// <param name="Fees">GP / SP fee parts (B1–B8).</param>
/// <param name="Member">The member panel.</param>
/// <param name="Worksheet">The stored worksheet.</param>
/// <param name="Bordx">SP_Claim_Bordx rows.</param>
/// <param name="Warnings">Member look-up warnings.</param>
public record InvoiceEnquiryCase(Dictionary<string, object?> Case, Dictionary<string, object?>? Fees, InvoiceMember? Member, StoredWorksheet Worksheet,
    List<Dictionary<string, object?>> Bordx, List<string> Warnings);

/// <summary>A document of the case (scanned / guarantee letter folders).</summary>
/// <param name="Group">Guarantee Letter, SP_ScannedDoc, OtherScannedDoc or Reim_ScannedDoc.</param>
/// <param name="FileName">File name.</param>
public record CaseDocument(string Group, string FileName);

/// <summary>Data access of the enquiry (<c>cis.InvEnq_*</c>).</summary>
public interface IInvoiceEnquiryRepository
{
    /// <summary>SP_INV_CLAIM and the fee parts.</summary>
    Task<(Dictionary<string, object?>? Case, Dictionary<string, object?>? Fees)> CaseAsync(IDbSession session, string casNumber);

    /// <summary>The next / previous case.</summary>
    Task<string?> StepAsync(IDbSession session, string casNumber, bool next);

    /// <summary>Claim and the seven worksheet tables (Actual, ClinicLmt, App, Limit, Member, Clinic, ClinicApp).</summary>
    Task<List<Dictionary<string, object?>?>> WorksheetAsync(IDbSession session, string casNumber);

    /// <summary>SP_Claim_Bordx.</summary>
    Task<List<Dictionary<string, object?>>> BordxAsync(IDbSession session, string casNumber);

    /// <summary>Payments (clinic and member) and receipts.</summary>
    Task<(List<Dictionary<string, object?>> Payments, List<Dictionary<string, object?>> Receipts)> PaymentsAsync(IDbSession session, string casNumber);

    /// <summary>Medication requests and their medicines.</summary>
    Task<(List<Dictionary<string, object?>> Requests, List<Dictionary<string, object?>> Medicines)> MedRequestsAsync(IDbSession session, string casNumber);
}

/// <summary>Invoice > Clinic > Enquiry.</summary>
public interface IInvoiceEnquiryService
{
    /// <summary>CmdGo_Click.</summary>
    Task<InvoiceEnquiryCase> FindAsync(string? casNumber);

    /// <summary>CmdNext_Click / CmdPrev_Click: the next / previous case number.</summary>
    Task<string> StepAsync(string? casNumber, bool next);

    /// <summary>Payment tab.</summary>
    Task<(List<Dictionary<string, object?>> Payments, List<Dictionary<string, object?>> Receipts)> PaymentsAsync(string? casNumber);

    /// <summary>Medication request tab.</summary>
    Task<(List<Dictionary<string, object?>> Requests, List<Dictionary<string, object?>> Medicines)> MedRequestsAsync(string? casNumber);

    /// <summary>Documents tab (null = the Outpatient folder is not configured).</summary>
    List<CaseDocument>? Documents(string? casNumber);

    /// <summary>A document's content.</summary>
    Task<(byte[] Content, string FileName)> DocumentAsync(string? casNumber, string? group, string? fileName);
}

/// <inheritdoc />
public sealed partial class InvoiceEnquiryService(IDbSessions sessions, IInvoiceEnquiryRepository repository, IClinicInvoiceService invoices,
    IWorksheetRepository worksheet, IFileDrop files) : IInvoiceEnquiryService
{
    /// <summary>Desktop texts.</summary>
    public const string NotFound = "Not Record Found ! ", Last = "Last Record Reached! ", First = "First Record Reached! \nNot a valid Case Number! ";

    /// <summary>Guarantee letter file types (addTheFiles).</summary>
    private static readonly string[] LetterTypes = [".jpg", ".tif", ".gif", ".doc", ".pdf", ".png", ".msg"];

    /// <summary>The worksheet slot columns in grid order.</summary>
    private static readonly string[] Slots =
    [
        "CLMGPFeeBasicMed", "CLMGPFeeMedInj", "CLMGPFeeLTMed", "CLMSPFeeBasicMed", "CLMSPFeeLTMed", "CLMSPFeeMedInj", "CLMXRay", "CLMUltraSound", "CLMECG",
        "CLMDressing", "CLMNebulizer", "CLMID", "CLMTSSTO", "CLMMinorSurgery", "CLMPOthers", "CLMBloodDrawing", "CLMBloodGlucose", "CLMUrine", "CLMPapSmear",
        "CLMLOthers", "CLMDTPPolio", "CLMDTPPolioHIB", "CLMHepatitisB", "CLMMMR", "CLMIOthers", "CLMDental", "CLMMaternity", "CLMPRVCheckup", "CLMVision",
        "CLMTyphiadJabNoDed", "CLMOthersNoDed",
    ];

    /// <inheritdoc />
    public async Task<InvoiceEnquiryCase> FindAsync(string? casNumber)
    {
        var cas = Case(casNumber);
        Dictionary<string, object?>? row, fees;
        List<Dictionary<string, object?>?> tables;
        List<Dictionary<string, object?>> bordx;
        WorksheetLookups lookups;
        await using (var s = await sessions.OpenAsync())
        {
            (row, fees) = await repository.CaseAsync(s, cas);
            if (row is null) throw new NotFoundException(NotFound);
            tables = await repository.WorksheetAsync(s, cas);
            bordx = await repository.BordxAsync(s, cas);
            lookups = await worksheet.LookupsAsync(s, new WorksheetInput { ClinicCode = S(row, "ClinicCode") });
        }
        var warnings = new List<string>();
        InvoiceMember? member = null;
        try
        {
            member = await invoices.MemberAsync(S(row, "MBMnumber"), S(row, "CoverID"));
            warnings.AddRange(member.Warnings);
        }
        catch (BusinessRuleException e)
        {
            warnings.Add(e.Message);
        }
        return new InvoiceEnquiryCase(row, fees, member, Stored(lookups.Categories, tables), bordx, warnings);
    }

    /// <summary>WSInfo: the seven tables as the worksheet's columns.</summary>
    public static StoredWorksheet Stored(IReadOnlyList<(string Code, string Description)> categories, List<Dictionary<string, object?>?> t)
    {
        Dictionary<string, object?>? T(int i) => i < t.Count ? t[i] : null;
        decimal? V(int table, int slot) => T(table)?.GetValueOrDefault(Slots[slot]) is { } o && o is not DBNull ? Convert.ToDecimal(o, CultureInfo.InvariantCulture) : null;
        var rows = categories.Select((c, i) => i < Slots.Length
            ? new WorksheetRow(c.Code, c.Description, V(1, i), V(2, i), V(7, i), V(6, i), V(4, i), V(3, i), V(5, i))
            : new WorksheetRow(c.Code, c.Description, null, null, null, null, null, null, null)).ToList();
        return new StoredWorksheet(rows, T(0));
    }

    /// <inheritdoc />
    public async Task<string> StepAsync(string? casNumber, bool next)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.StepAsync(s, (casNumber ?? "").Trim().ToUpperInvariant(), next) ?? throw new NotFoundException(next ? Last : First);
    }

    /// <inheritdoc />
    public async Task<(List<Dictionary<string, object?>> Payments, List<Dictionary<string, object?>> Receipts)> PaymentsAsync(string? casNumber)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.PaymentsAsync(s, Case(casNumber));
    }

    /// <inheritdoc />
    public async Task<(List<Dictionary<string, object?>> Requests, List<Dictionary<string, object?>> Medicines)> MedRequestsAsync(string? casNumber)
    {
        await using var s = await sessions.OpenAsync();
        return await repository.MedRequestsAsync(s, Case(casNumber));
    }

    /// <inheritdoc />
    public List<CaseDocument>? Documents(string? casNumber)
    {
        var cas = Case(casNumber);
        if (files.FolderOf(FileFolders.Outpatient) is not { Length: > 0 } root) return null;
        var list = new List<CaseDocument>();
        void Folder(string group, Func<string, bool> filter)
        {
            var dir = Path.Combine(root, group, cas);
            if (Directory.Exists(dir))
                list.AddRange(Directory.EnumerateFiles(dir).Select(Path.GetFileName).OfType<string>().Where(filter).Order(StringComparer.OrdinalIgnoreCase)
                    .Select(f => new CaseDocument(group, f)));
        }
        void Single(string group)
        {
            foreach (var ext in new[] { ".pdf", ".jpg" })
                if (File.Exists(Path.Combine(root, group, cas + ext))) list.Add(new CaseDocument(group, cas + ext));
        }
        Folder("Guarantee Letter", f => LetterTypes.Contains(Path.GetExtension(f), StringComparer.OrdinalIgnoreCase));
        Single("SP_ScannedDoc");
        Folder("OtherScannedDoc", _ => true);
        Folder("Reim_ScannedDoc", _ => true);
        Single("Reim_ScannedDoc");
        return list;
    }

    /// <inheritdoc />
    public async Task<(byte[] Content, string FileName)> DocumentAsync(string? casNumber, string? group, string? fileName)
    {
        var cas = Case(casNumber);
        var list = Documents(cas) ?? throw new NotConfiguredException("The outpatient documents folder is not configured (Files:Folders:Outpatient).");
        var doc = list.FirstOrDefault(d => string.Equals(d.Group, group, StringComparison.OrdinalIgnoreCase)
            && string.Equals(d.FileName, fileName, StringComparison.OrdinalIgnoreCase)) ?? throw new NotFoundException("File not found.");
        var root = files.FolderOf(FileFolders.Outpatient)!;
        var path = File.Exists(Path.Combine(root, doc.Group, cas, doc.FileName)) ? Path.Combine(root, doc.Group, cas, doc.FileName) : Path.Combine(root, doc.Group, doc.FileName);
        return (await File.ReadAllBytesAsync(path), doc.FileName);
    }

    /// <summary>A case number (letters and digits only, so it cannot leave the document folders).</summary>
    private static string Case(string? casNumber)
    {
        var cas = (casNumber ?? "").Trim().ToUpperInvariant();
        return CaseNo().IsMatch(cas) ? cas : throw new NotFoundException(NotFound);
    }

    private static string S(Dictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";

    [GeneratedRegex("^[A-Z0-9]{1,20}$")]
    private static partial Regex CaseNo();
}
