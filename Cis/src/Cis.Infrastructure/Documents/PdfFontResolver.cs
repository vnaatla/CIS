using PdfSharp.Fonts;

namespace Cis.Infrastructure.Documents;

/// <summary>
/// PDFsharp font source (same approach as GENESIS). Sans-serif families map to Arial / DejaVu Sans / Liberation Sans and monospaced ones to
/// Courier New / DejaVu Sans Mono / Liberation Mono, read from the Windows and Linux font folders. Missing bold files are simulated.
/// </summary>
public sealed class PdfFontResolver : IFontResolver
{
    private const string SansFace = "Cis-Sans";
    private const string SansBoldFace = "Cis-Sans-Bold";
    private const string MonoFace = "Cis-Mono";
    private const string MonoBoldFace = "Cis-Mono-Bold";

    /// <summary>Folders searched, recursively, for TrueType files.</summary>
    public static readonly IReadOnlyList<string> DefaultFontFolders =
    [
        Environment.GetFolderPath(Environment.SpecialFolder.Fonts),
        "/usr/share/fonts", "/usr/local/share/fonts", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts"),
    ];

    private static readonly IReadOnlySet<string> MonospacedFamilies = new HashSet<string>(
        ["Courier New", "Courier", "Consolas", "Lucida Console", "Monospace", "DejaVu Sans Mono", "Liberation Mono"], StringComparer.OrdinalIgnoreCase);

    private static readonly IReadOnlyDictionary<string, string[]> Candidates = new Dictionary<string, string[]>
    {
        [SansFace] = ["arial.ttf", "DejaVuSans.ttf", "LiberationSans-Regular.ttf", "NotoSans-Regular.ttf", "FreeSans.ttf"],
        [SansBoldFace] = ["arialbd.ttf", "DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf", "NotoSans-Bold.ttf", "FreeSansBold.ttf"],
        [MonoFace] = ["cour.ttf", "DejaVuSansMono.ttf", "LiberationMono-Regular.ttf", "NotoSansMono-Regular.ttf", "FreeMono.ttf"],
        [MonoBoldFace] = ["courbd.ttf", "DejaVuSansMono-Bold.ttf", "LiberationMono-Bold.ttf", "NotoSansMono-Bold.ttf", "FreeMonoBold.ttf"],
    };

    private readonly Dictionary<string, string> _files = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Indexes the TrueType files under <paramref name="fontFolders"/> (the system folders by default).</summary>
    public PdfFontResolver(IEnumerable<string>? fontFolders = null)
    {
        foreach (var folder in (fontFolders ?? DefaultFontFolders).Where(f => !string.IsNullOrEmpty(f) && Directory.Exists(f)))
            foreach (var file in Directory.EnumerateFiles(folder, "*.ttf", SearchOption.AllDirectories))
                _files.TryAdd(Path.GetFileName(file), file);
    }

    /// <summary>Whether a regular sans-serif and monospaced font were both found.</summary>
    public bool HasFonts => Find(SansFace) is not null && Find(MonoFace) is not null;

    /// <inheritdoc />
    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var mono = MonospacedFamilies.Contains(familyName);
        var regular = mono ? MonoFace : SansFace;
        var bold = mono ? MonoBoldFace : SansBoldFace;
        if (isBold && Find(bold) is not null) return new FontResolverInfo(bold, false, isItalic);
        return Find(regular) is null ? null : new FontResolverInfo(regular, isBold, isItalic);
    }

    /// <inheritdoc />
    public byte[]? GetFont(string faceName) => Find(faceName) is { } path ? File.ReadAllBytes(path) : null;

    private string? Find(string faceName) =>
        Candidates.TryGetValue(faceName, out var names) ? names.Select(n => _files.GetValueOrDefault(n)).FirstOrDefault(p => p is not null) : null;
}
