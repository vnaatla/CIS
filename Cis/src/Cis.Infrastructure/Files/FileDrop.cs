using Cis.Application.Files;
using Microsoft.Extensions.Options;

namespace Cis.Infrastructure.Files;

/// <summary>
/// Server folders the desktop wrote to (configuration section <c>Files</c>, keys = the SVRSuper.ini settings they replace). Empty =
/// files are only downloaded by the browser, not copied to a share.
/// </summary>
public class FileDropOptions
{
    /// <summary>Configuration section name.</summary>
    public const string SectionName = "Files";

    /// <summary>Folder per kind of file, e.g. <c>ClaimAscii</c> → <c>\\svr-doc\ClinicalClaimASCII</c>.</summary>
    public Dictionary<string, string> Folders { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

/// <summary><see cref="IFileDrop"/> on the folders of <see cref="FileDropOptions"/>; file names that would leave the folder are refused.</summary>
public sealed class FileDrop(IOptions<FileDropOptions> options) : IFileDrop
{
    /// <inheritdoc />
    public string? FolderOf(string folder) =>
        options.Value.Folders.TryGetValue(folder, out var path) && !string.IsNullOrWhiteSpace(path) ? path : null;

    /// <inheritdoc />
    public async Task<string?> SaveAsync(string folder, string fileName, byte[] content)
    {
        if (FolderOf(folder) is not { } root) return null;
        var name = Path.GetFileName(fileName);
        var fullRoot = Path.GetFullPath(root);
        var path = Path.GetFullPath(Path.Combine(fullRoot, name));
        if (name.Trim('.', ' ').Length == 0 || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 ||
            !string.Equals(Path.GetDirectoryName(path), fullRoot.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException($"Invalid file name '{fileName}'.", nameof(fileName));
        Directory.CreateDirectory(fullRoot);
        await File.WriteAllBytesAsync(path, content);
        return path;
    }
}
