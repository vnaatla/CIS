namespace Cis.Application.Documents;

/// <summary>How a report cell is formatted.</summary>
public enum ReportFormat
{
    /// <summary>Plain text.</summary>
    Text,
    /// <summary>Whole number.</summary>
    Integer,
    /// <summary>Number with two decimals and thousands separators (RM amounts).</summary>
    Money,
    /// <summary>Date shown as dd/mm/yyyy.</summary>
    Date,
    /// <summary>Number as stored.</summary>
    Number,
}

/// <summary>Output file types of the report engine.</summary>
public enum ReportFileType
{
    /// <summary>Excel workbook (.xlsx), replaces the desktop's ExportToExcel* routines and .xlt templates.</summary>
    Excel,
    /// <summary>PDF, replaces the Crystal Reports and DataReport print-outs.</summary>
    Pdf,
    /// <summary>Comma-separated text.</summary>
    Csv,
}

/// <summary>One column of a report.</summary>
/// <param name="Key">Row dictionary key (case-insensitive), usually the SQL column name.</param>
/// <param name="Header">Column caption (the desktop's ListView / Excel header).</param>
/// <param name="Format">Cell format.</param>
/// <param name="Total">Sum the column in a totals row.</param>
public record ReportColumn(string Key, string Header, ReportFormat Format = ReportFormat.Text, bool Total = false);

/// <summary>A tabular report ready to be written as Excel, PDF or CSV.</summary>
/// <param name="Title">Report title (first line).</param>
/// <param name="Columns">Columns in display order.</param>
/// <param name="Rows">Data rows (case-insensitive keys).</param>
public record ReportDocument(string Title, IReadOnlyList<ReportColumn> Columns, IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows)
{
    /// <summary>Lines under the title (filters used, run date ...).</summary>
    public IReadOnlyList<string> Subtitles { get; init; } = [];

    /// <summary>Base file name without extension.</summary>
    public string FileName { get; init; } = "Report";

    /// <summary>
    /// Optional legacy template name (e.g. <c>CISTemplate\ClaimReport</c>). When <see cref="IReportTemplateStore"/> finds it the Excel
    /// writer fills the template from <see cref="TemplateStartRow"/>; otherwise the layout is generated.
    /// </summary>
    public string? Template { get; init; }

    /// <summary>First data row in the template (1-based).</summary>
    public int TemplateStartRow { get; init; } = 2;

    /// <summary>Print the PDF in landscape (wide listings).</summary>
    public bool Landscape { get; init; } = true;
}

/// <summary>A generated file.</summary>
/// <param name="Content">File bytes.</param>
/// <param name="ContentType">MIME type.</param>
/// <param name="FileName">Download file name with extension.</param>
public record FileOutput(byte[] Content, string ContentType, string FileName);
