using System.Globalization;
using Cis.Application.Abstractions;
using Cis.Application.Documents;
using Cis.Domain.Errors;

namespace Cis.Application.Invoice;

/// <summary>The MMA schedule filters (CmdSearch_Click).</summary>
/// <param name="MainBp">Main body part.</param>
/// <param name="Opcs">OPCS code.</param>
/// <param name="Procedure">Procedure (narrative).</param>
/// <param name="SurgCat">Surgical category.</param>
/// <param name="AnaesCat">Anaesthetic category.</param>
/// <param name="Part">MMA part.</param>
/// <param name="PageFrom">Page from.</param>
/// <param name="PageTo">Page to.</param>
/// <param name="CptCode">CPT code.</param>
/// <param name="CptDesc">CPT description.</param>
public record MmaFilters(string? MainBp, string? Opcs, string? Procedure, string? SurgCat, string? AnaesCat, string? Part, string? PageFrom, string? PageTo,
    string? CptCode, string? CptDesc);

/// <summary>A row of lstMMAListing.</summary>
/// <param name="MainBp">Main body part.</param>
/// <param name="Opcs">OPCS code.</param>
/// <param name="Narrative">Procedure.</param>
/// <param name="SurgCat">Surgical category (without its first character).</param>
/// <param name="AnaesCat">Anaesthetic category (without its first character).</param>
/// <param name="SurgeonFee">Surgeon fee (13th, not part A).</param>
/// <param name="AnaesFee">Anaesthetist fee (13th, not part A).</param>
/// <param name="PartAFee">Part A cost (part A only).</param>
/// <param name="Page">Page.</param>
/// <param name="Part">MMA part.</param>
/// <param name="CptCode">CPT code.</param>
/// <param name="CptDesc">CPT description.</param>
public record MmaRow(string MainBp, string Opcs, string Narrative, string SurgCat, string AnaesCat, string SurgeonFee, string AnaesFee, string PartAFee,
    string Page, string Part, string CptCode, string CptDesc);

/// <summary>Data access of the MMA schedule.</summary>
public interface IMmaScheduleRepository
{
    /// <summary>VIEW_MMA_Schedule rows (filters already reduced).</summary>
    Task<List<Dictionary<string, object?>>> SearchAsync(IDbSession session, MmaFilters filters);
}

/// <summary>Invoice > Other > MMA/CPT.</summary>
public interface IMmaScheduleService
{
    /// <summary>CmdSearch_Click.</summary>
    Task<List<MmaRow>> SearchAsync(MmaFilters filters);

    /// <summary>CmdPrint_Click.</summary>
    Task<FileOutput> ExportAsync(MmaFilters filters);
}

/// <inheritdoc />
public sealed class MmaScheduleService(IDbSessions sessions, IMmaScheduleRepository repository, IReportExporter exporter) : IMmaScheduleService
{
    /// <summary>Desktop text.</summary>
    public const string NotFound = "Record not found";

    /// <summary>The list's columns (export).</summary>
    public static readonly IReadOnlyList<ReportColumn> Columns =
    [
        new("MainBp", "Main Body Part"), new("Opcs", "OPCS Code"), new("Narrative", "Procedure"), new("SurgCat", "Surgical Cat"),
        new("AnaesCat", "Anaesthetic Cat"), new("SurgeonFee", "Surgeon Fee"), new("AnaesFee", "Anaesthetist Fee"), new("PartAFee", "Part A Fee"),
        new("Page", "Page"), new("Part", "Part"), new("CptCode", "CPT Code"), new("CptDesc", "CPT Description"),
    ];

    /// <inheritdoc />
    public async Task<List<MmaRow>> SearchAsync(MmaFilters filters)
    {
        await using var s = await sessions.OpenAsync();
        var rows = (await repository.SearchAsync(s, Reduce(filters))).Select(Row).ToList();
        return rows.Count == 0 ? throw new NotFoundException(NotFound) : rows;
    }

    /// <inheritdoc />
    public async Task<FileOutput> ExportAsync(MmaFilters filters)
    {
        var rows = await SearchAsync(filters);
        var dicts = rows.Select(r => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["MainBp"] = r.MainBp, ["Opcs"] = r.Opcs, ["Narrative"] = r.Narrative, ["SurgCat"] = r.SurgCat, ["AnaesCat"] = r.AnaesCat,
            ["SurgeonFee"] = r.SurgeonFee, ["AnaesFee"] = r.AnaesFee, ["PartAFee"] = r.PartAFee, ["Page"] = r.Page, ["Part"] = r.Part,
            ["CptCode"] = r.CptCode, ["CptDesc"] = r.CptDesc,
        }).ToList();
        return exporter.Export(new ReportDocument("MMA Schedule", Columns, dicts) { FileName = "MMASchedule" }, ReportFileType.Excel);
    }

    /// <summary>CmdSearch_Click's criteria: ChkStr (upper, trimmed) where the desktop used it; the categories without their first character.</summary>
    public static MmaFilters Reduce(MmaFilters f)
    {
        static string? T(string? v) => string.IsNullOrWhiteSpace(v) ? null : v.Trim();
        static string? U(string? v) => T(v)?.ToUpperInvariant();
        // Mid(x, 2, 8) of the untrimmed text: characters 2–9; empty when the text has one character.
        static string? Mid(string? v) => string.IsNullOrEmpty(v) || v.Trim().Length == 0 ? null : v.Length < 2 ? "" : v.Substring(1, Math.Min(8, v.Length - 1));
        return new MmaFilters(U(f.MainBp), string.IsNullOrEmpty(f.Opcs) ? null : f.Opcs.Trim().ToUpperInvariant(), U(f.Procedure), Mid(f.SurgCat),
            Mid(f.AnaesCat), U(f.Part), T(f.PageFrom), T(f.PageTo), string.IsNullOrEmpty(f.CptCode) ? null : f.CptCode.Trim(), T(f.CptDesc));
    }

    /// <summary>MMAListing: categories without their first character; surgeon / anaesthetist fees unless part A, part A cost for part A.</summary>
    public static MmaRow Row(Dictionary<string, object?> r)
    {
        var part = S(r, "ProcPart");
        static string Mid(string v) => v.Length < 2 ? "" : v.Substring(1, Math.Min(8, v.Length - 1));
        return new MmaRow(S(r, "ProcMainBPCode"), S(r, "OPCSCode"), S(r, "OPCSNarrative"), Mid(S(r, "ProcSurgCat")), Mid(S(r, "ProcAnaesCat")),
            part != "A" ? S(r, "Surgeon13thNew") : "", part != "A" ? S(r, "Anesthetist13thNew") : "", part == "A" ? S(r, "PartACost") : "",
            S(r, "ProcPageNo"), part, S(r, "CPTCode"), S(r, "CPTDesc"));
    }

    private static string S(Dictionary<string, object?> r, string k) =>
        r.TryGetValue(k, out var v) && v is not null ? Convert.ToString(v, CultureInfo.InvariantCulture)!.Trim() : "";
}
