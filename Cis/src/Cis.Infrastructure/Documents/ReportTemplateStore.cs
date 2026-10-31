using Cis.Application.Documents;
using Microsoft.Extensions.Options;

namespace Cis.Infrastructure.Documents;

/// <summary>Settings of the report engine (configuration section <c>Reports</c>).</summary>
public class ReportOptions
{
    /// <summary>Configuration section name.</summary>
    public const string SectionName = "Reports";

    /// <summary>
    /// Folder holding copies of the desktop's <c>\\svr-doc\HISREPORTS</c> templates converted to .xltx / .xlsx (same sub-folders,
    /// e.g. <c>CISTemplate\ClaimReport.xltx</c>). Empty = every layout is generated in code.
    /// </summary>
    public string? TemplateRoot { get; set; }

    /// <summary>Most rows a report returns to the screen or a file (protects the server; the desktop had no limit).</summary>
    public int MaxRows { get; set; } = 100_000;
}

/// <summary><see cref="IReportTemplateStore"/> on a folder (<see cref="ReportOptions.TemplateRoot"/>); paths outside it are refused.</summary>
public sealed class ReportTemplateStore(IOptions<ReportOptions> options) : IReportTemplateStore
{
    private static readonly string[] Extensions = [".xltx", ".xlsx"];

    /// <inheritdoc />
    public string? Find(string template)
    {
        var root = options.Value.TemplateRoot;
        if (string.IsNullOrWhiteSpace(root) || string.IsNullOrWhiteSpace(template)) return null;
        var relative = Path.ChangeExtension(template.Replace('\\', Path.DirectorySeparatorChar).TrimStart(Path.DirectorySeparatorChar), null);
        var fullRoot = Path.GetFullPath(root);
        foreach (var ext in Extensions)
        {
            var path = Path.GetFullPath(Path.Combine(fullRoot, relative + ext));
            if (path.StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase) && File.Exists(path)) return path;
        }
        return null;
    }
}
