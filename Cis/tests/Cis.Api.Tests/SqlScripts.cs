using System.Text;
using System.Text.RegularExpressions;
using Dapper;
using Microsoft.Data.SqlClient;

namespace Cis.Api.Tests;

/// <summary>Builds sql/cis-main.sql from sql/main/*.sql and deploys it to the test database.</summary>
public static class SqlScripts
{
    public static string CombinedPath => Path.Combine(RepoPaths.CisRoot, "sql", "cis-main.sql");

    private static readonly SemaphoreSlim Gate = new(1, 1);
    private static bool _deployed;

    public static string BuildCombined()
    {
        var sb = new StringBuilder();
        sb.Append("""
            -- GENERATED FILE: built from sql/main/*.sql. Edit those files, then rebuild this one with
            --     $env:CIS_WRITE_SQL='1'; dotnet test --filter Combined_script_is_up_to_date
            -- How to run: sqlcmd -S <server> -E -d CISDB -i cis-main.sql   (or open it in SSMS on CISDB and press F5).
            -- Check the database names at the top of the first section first.
            -- Safe to run again: it only creates or replaces cis.* procedures and synonyms; tables and data are not changed.

            """);
        foreach (var file in Directory.GetFiles(Path.Combine(RepoPaths.CisRoot, "sql", "main"), "*.sql").Order(StringComparer.Ordinal))
            sb.Append($"\n-- ===== main/{Path.GetFileName(file)} =====\n").Append(Normalize(File.ReadAllText(file)).TrimEnd('\n')).Append('\n');
        sb.Append("\nPRINT N'Done.';\nGO\nSET NOEXEC OFF;\nGO\n");
        return Normalize(sb.ToString());
    }

    public static string Normalize(string s) => s.Replace("\r\n", "\n");

    /// <summary>Runs the combined script once per test run (it is idempotent).</summary>
    public static async Task EnsureDeployedAsync(string connectionString)
    {
        await Gate.WaitAsync();
        try
        {
            if (_deployed) return;
            await using var conn = new SqlConnection(connectionString);
            await conn.OpenAsync();
            foreach (var batch in Regex.Split(File.ReadAllText(CombinedPath), @"^\s*GO\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase))
                if (!string.IsNullOrWhiteSpace(batch)) await conn.ExecuteAsync(batch, commandTimeout: 120);
            _deployed = true;
        }
        finally
        {
            Gate.Release();
        }
    }
}

public class SqlScriptTests
{
    [Fact]
    public void Combined_script_is_up_to_date()
    {
        var expected = SqlScripts.BuildCombined();
        if (Environment.GetEnvironmentVariable("CIS_WRITE_SQL") == "1") File.WriteAllText(SqlScripts.CombinedPath, expected);
        Assert.True(File.Exists(SqlScripts.CombinedPath) && expected == SqlScripts.Normalize(File.ReadAllText(SqlScripts.CombinedPath)),
            "sql/cis-main.sql is missing or out of date; rebuild it with CIS_WRITE_SQL=1 dotnet test --filter Combined_script_is_up_to_date");
    }

    [Fact]
    public void Every_script_only_creates_cis_objects()
    {
        foreach (var file in Directory.GetFiles(Path.Combine(RepoPaths.CisRoot, "sql", "main"), "*.sql"))
        {
            var text = File.ReadAllText(file);
            foreach (Match m in Regex.Matches(text, @"CREATE\s+OR\s+ALTER\s+(PROCEDURE|FUNCTION|VIEW)\s+([\w\.\[\]]+)", RegexOptions.IgnoreCase))
                Assert.StartsWith("cis.", m.Groups[2].Value, StringComparison.OrdinalIgnoreCase);
            // Temp tables (#...) inside copied legacy procedures are allowed; real tables are never altered, dropped or truncated.
            Assert.DoesNotMatch(new Regex(@"\b(ALTER|DROP)\s+TABLE\s+(?!#)|\bTRUNCATE\b", RegexOptions.IgnoreCase), text);
        }
    }

    [Fact]
    public void Scripts_avoid_functions_newer_than_compatibility_level_120()
    {
        var banned = new Regex(@"\b(STRING_AGG|STRING_SPLIT|CONCAT_WS|TRIM|TRANSLATE|GREATEST|LEAST)\s*\(", RegexOptions.IgnoreCase);
        foreach (var file in Directory.GetFiles(Path.Combine(RepoPaths.CisRoot, "sql", "main"), "*.sql"))
            Assert.False(banned.IsMatch(Regex.Replace(File.ReadAllText(file), @"\b[LR]TRIM\s*\(", "")), $"{Path.GetFileName(file)} uses a function CISDB (level 120) / CISWorksheet (level 100) does not have");
    }
}
