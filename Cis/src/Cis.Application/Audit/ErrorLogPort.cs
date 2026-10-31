namespace Cis.Application.Audit;

/// <summary>Writes unexpected server errors to the legacy error table (desktop <c>Medix.ErrorLog</c>).</summary>
public interface IErrorLog
{
    /// <summary>Records an error; never throws (a failing error log must not hide the original error).</summary>
    Task WriteAsync(string screen, string method, string description, string? userCode);
}
