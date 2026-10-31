using System.Data;
using Cis.Api.Data;
using Dapper;

namespace Cis.Api.Audit;

/// <summary>Writes unexpected server errors to the legacy error table (desktop <c>Medix.ErrorLog</c>).</summary>
public interface IErrorLog
{
    /// <summary>Records an error; never throws (a failing error log must not hide the original error).</summary>
    Task WriteAsync(string screen, string method, string description, string? userCode);
}

/// <summary>
/// <see cref="IErrorLog"/> on <c>HISMaintenance.dbo.Tbl_ErrorLog</c> (SysName 'CIS') through <c>cis.ErrorLog_Write</c>. Runs on its own
/// connection, outside the request's (rolled back) transaction. The desktop also decremented USR.CIScnt and e-mailed IT here because the
/// error closed the application; the web session survives an error, so neither is done.
/// </summary>
public sealed class ErrorLog(DbConnections db, ILogger<ErrorLog> log) : IErrorLog
{
    /// <summary>Longest FormName / MethodName the table accepts.</summary>
    public const int NameLength = 50;

    /// <inheritdoc />
    public async Task WriteAsync(string screen, string method, string description, string? userCode)
    {
        try
        {
            await using var conn = db.Main();
            await conn.ExecuteAsync("cis.ErrorLog_Write",
                new { formName = Cut(screen), methodName = Cut(method), description = $"{description} (user {userCode ?? "anonymous"})" },
                commandType: CommandType.StoredProcedure);
        }
        catch (Exception e)
        {
            log.LogWarning(e, "Could not write Tbl_ErrorLog");
        }
    }

    private static string Cut(string text) => text.Length > NameLength ? text[..NameLength] : text;
}
