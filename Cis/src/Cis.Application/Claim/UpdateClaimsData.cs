using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Domain.Errors;

namespace Cis.Application.Claim;

/// <summary>The four uploads of frmUpdateClaimsData.</summary>
public enum ClaimDataKind
{
    /// <summary>FFS (CLMMCOProcFee).</summary>
    Ffs,
    /// <summary>Co-pay collected (CLMCoPayColl).</summary>
    Copay,
    /// <summary>Bordereaux amount (CLMBordxAmt).</summary>
    BordxAmt,
    /// <summary>Bordereaux date (Claim month / year, Bordx.BordxDate).</summary>
    BordxDate,
}

/// <summary>A row the verification rejected (lstErrorList).</summary>
/// <param name="Row">Sheet row.</param>
/// <param name="BordxNo">Bordx / case number.</param>
/// <param name="AppCode">Approval code.</param>
/// <param name="AppDate">Approval date.</param>
/// <param name="Amount">Amount.</param>
/// <param name="Date">Return / send date.</param>
public record ClaimDataError(int Row, string BordxNo, string AppCode, string AppDate, string Amount, string Date);

/// <summary>Verify / upload result.</summary>
/// <param name="Message">Desktop message.</param>
/// <param name="Rows">Rows read.</param>
/// <param name="Errors">Rejected rows.</param>
/// <param name="Updated">Records updated (uploads).</param>
public record ClaimDataResult(string Message, int Rows, List<ClaimDataError> Errors, int Updated = 0);

/// <summary>Data access of Update Claims Data (<c>cis.ClaimData_*</c>).</summary>
public interface IUpdateClaimsDataRepository
{
    /// <summary>One upload row; records updated.</summary>
    Task<int> UpdateAsync(IDbSession session, ClaimDataKind kind, string key, decimal? amount, DateTime? date, DateTime today, string userCode);
}

/// <summary>Claim Bordx > Update Claims Data.</summary>
public interface IUpdateClaimsDataService
{
    /// <summary>cmdVerify / cmdVerifyReturn / cmdVerifySend (and Bordx Date, which had no handler).</summary>
    ClaimDataResult Verify(ClaimDataKind kind, Stream file);

    /// <summary>cmdUpload / cmdReturnDate / cmdUploadSend / cmdUpdateBordxDate.</summary>
    Task<ClaimDataResult> UploadAsync(ClaimDataKind kind, Stream file, string userCode);
}

/// <inheritdoc />
public sealed class UpdateClaimsDataService(IDbSessions sessions, IUpdateClaimsDataRepository repository, ISpreadsheetReader reader, TimeProvider clock)
    : IUpdateClaimsDataService
{
    /// <summary>Desktop texts (spelling kept).</summary>
    public const string Verified = "File verified successfully!", Errors = "Error encoutered. Please check the file!",
        FileFirst = "Please select file to upload!", Uploaded = "Records have been uploaded!";

    /// <summary>Columns per upload: key, approval code, approval date, amount, date (0 = not read).</summary>
    public static (int Key, int Code, int AppDate, int Amount, int Date) Columns(ClaimDataKind kind) => kind switch
    {
        ClaimDataKind.Ffs => (1, 8, 9, 10, 0),
        ClaimDataKind.Copay => (2, 18, 19, 3, 20),
        ClaimDataKind.BordxAmt => (2, 0, 0, 3, 20),
        _ => (2, 0, 0, 0, 20),
    };

    /// <inheritdoc />
    public ClaimDataResult Verify(ClaimDataKind kind, Stream file)
    {
        var rows = Read(kind, file);
        var errors = rows.Select(r => Check(kind, r.Row, r.Cells)).OfType<ClaimDataError>().ToList();
        return new ClaimDataResult(errors.Count == 0 ? Verified : Errors, rows.Count, errors);
    }

    /// <inheritdoc />
    public async Task<ClaimDataResult> UploadAsync(ClaimDataKind kind, Stream file, string userCode)
    {
        var rows = Read(kind, file);
        // The desktop enabled the upload only after a clean verification; the API verifies the file again.
        var errors = rows.Select(r => Check(kind, r.Row, r.Cells)).OfType<ClaimDataError>().ToList();
        if (errors.Count > 0) return new ClaimDataResult(Errors, rows.Count, errors);
        var c = Columns(kind);
        var today = clock.GetLocalNow().Date;
        var updated = 0;
        await using var s = await sessions.BeginAsync();
        foreach (var (_, cells) in rows)
            updated += await repository.UpdateAsync(s, kind, cells[c.Key].Trim(), c.Amount > 0 ? Amount(cells[c.Amount]) : null,
                c.Date > 0 ? Date(cells[c.Date]) : null, today, userCode);
        await s.CommitAsync();
        return new ClaimDataResult(Uploaded, rows.Count, [], updated);
    }

    /// <summary>The desktop's row checks: bordx no. of 8 or 10 characters starting N or S, approval code of 12 characters (FFS), a valid
    /// approval date (FFS), a numeric amount and a valid return / send date.</summary>
    public static ClaimDataError? Check(ClaimDataKind kind, int row, string[] cells)
    {
        var c = Columns(kind);
        string V(int i) => i > 0 && i < cells.Length ? cells[i] ?? "" : "";
        var key = V(c.Key);
        var error = (key.Length != 8 && key.Length != 10) || !(key.StartsWith('N') || key.StartsWith('S'));
        if (kind == ClaimDataKind.Ffs) error |= V(c.Code).Length != 12 || Date(V(c.AppDate)) is null;
        if (c.Amount > 0) error |= Amount(V(c.Amount)) is null;
        if (c.Date > 0) error |= Date(V(c.Date)) is null;
        return error ? new ClaimDataError(row, key, V(c.Code), V(c.AppDate), V(c.Amount), V(c.Date)) : null;
    }

    private List<(int Row, string[] Cells)> Read(ClaimDataKind kind, Stream file)
    {
        if (file is null || (file.CanSeek && file.Length == 0)) throw new BusinessRuleException(FileFirst);
        try
        {
            return reader.Read(file, 20);
        }
        catch (Exception e) when (e is not BusinessRuleException)
        {
            throw new BusinessRuleException(Errors);
        }
    }

    private static decimal? Amount(string text) => decimal.TryParse(text?.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out var d) ? d : null;

    private static DateTime? Date(string text) =>
        DateTime.TryParseExact(text?.Trim(), ["dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd", "dd-MMM-yyyy", "d-MMM-yyyy"], CultureInfo.InvariantCulture,
            DateTimeStyles.None, out var d) ? d : null;
}
