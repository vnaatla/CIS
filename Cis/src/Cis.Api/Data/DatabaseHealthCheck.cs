using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Diagnostics.HealthChecks;

namespace Cis.Api.Data;

/// <summary>Healthy when CISDB answers and the <c>cis</c> scripts are deployed (GET /api/health).</summary>
public sealed class DatabaseHealthCheck(DbConnections db) : IHealthCheck
{
    /// <inheritdoc />
    public async Task<HealthCheckResult> CheckHealthAsync(HealthCheckContext context, CancellationToken cancellationToken = default)
    {
        try
        {
            await using var conn = db.Main();
            await conn.OpenAsync(cancellationToken);
            await using var cmd = new SqlCommand("SELECT CASE WHEN OBJECT_ID(N'cis.Login_GetActiveUser', N'P') IS NULL THEN 0 ELSE 1 END", conn);
            var deployed = (int)(await cmd.ExecuteScalarAsync(cancellationToken) ?? 0) == 1;
            return deployed
                ? HealthCheckResult.Healthy("CISDB reachable, cis scripts deployed.")
                : HealthCheckResult.Degraded("CISDB reachable, but Cis/sql/cis-main.sql has not been run.");
        }
        catch (Exception e) when (e is SqlException or InvalidOperationException)
        {
            return HealthCheckResult.Unhealthy("CISDB not reachable.", e);
        }
    }
}
