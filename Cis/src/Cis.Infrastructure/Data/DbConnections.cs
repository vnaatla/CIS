using Cis.Domain.Errors;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace Cis.Infrastructure.Data;

/// <summary>
/// Creates connections from the <c>ConnectionStrings</c> section. Everything runs on <see cref="Main"/> (CISDB); the other legacy
/// databases (HISDB, HISMaintenance, CISWorksheet, CISPayment, HISWorkSheet, HISPayment, HISDMArchive) are reached through the
/// <c>cis.*</c> synonyms created by <c>sql/main/000_setup.sql</c>, so a save that touches several databases is one transaction.
/// External systems that are not on that server (SUN Systems, the Jupiter reporting replica) have their own optional strings.
/// </summary>
/// <param name="configuration">Application configuration.</param>
public class DbConnections(IConfiguration configuration)
{
    /// <summary>Name of the main connection string (CISDB).</summary>
    public const string MainName = "CisDb";

    /// <summary>Optional connection string of the SUN Systems accounting database (desktop: MediConnSUNSystem.UDL).</summary>
    public const string SunSystemsName = "SunSystems";

    /// <summary>Optional connection string of the Jupiter / Neptune reporting replica; the main database when not set.</summary>
    public const string ReportingName = "Reporting";

    /// <summary>A new, unopened connection to the main database (CISDB).</summary>
    public SqlConnection Main() => new(Required(MainName));

    /// <summary>A new, unopened connection to the reporting replica (the "Jupiter" menu items); falls back to the main database.</summary>
    public SqlConnection Reporting() => new(configuration.GetConnectionString(ReportingName) is { Length: > 0 } cs ? cs : Required(MainName));

    /// <summary>A new, unopened connection to the SUN Systems database, or a <see cref="NotConfiguredException"/> when it is not set up.</summary>
    public SqlConnection SunSystems() =>
        configuration.GetConnectionString(SunSystemsName) is { Length: > 0 } cs
            ? new SqlConnection(cs)
            : throw new NotConfiguredException("SUN Systems database not configured. Set ConnectionStrings:SunSystems on the API server.");

    /// <summary>True when the named optional connection string is set.</summary>
    public bool IsConfigured(string name) => !string.IsNullOrWhiteSpace(configuration.GetConnectionString(name));

    private string Required(string name) =>
        configuration.GetConnectionString(name) is { Length: > 0 } cs
            ? cs
            : throw new InvalidOperationException($"Connection string '{name}' is not configured.");
}
