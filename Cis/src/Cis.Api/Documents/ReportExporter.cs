using System.Text;

namespace Cis.Api.Documents;

/// <summary>Writes reports in the format the user asked for.</summary>
public interface IReportExporter
{
    /// <summary>The report as a downloadable file.</summary>
    FileOutput Export(ReportDocument report, ReportFileType type);
}

/// <inheritdoc />
public sealed class ReportExporter(ExcelReportWriter excel, PdfReportWriter pdf) : IReportExporter
{
    /// <summary>MIME type of CSV.</summary>
    public const string CsvContentType = "text/csv";

    /// <inheritdoc />
    public FileOutput Export(ReportDocument report, ReportFileType type) => type switch
    {
        ReportFileType.Excel => new FileOutput(excel.Write(report), ExcelReportWriter.ContentType, report.FileName + ".xlsx"),
        ReportFileType.Pdf => new FileOutput(pdf.Write(report), PdfReportWriter.ContentType, report.FileName + ".pdf"),
        _ => new FileOutput(Csv(report), CsvContentType, report.FileName + ".csv"),
    };

    /// <summary>CSV with a UTF-8 BOM (opens in Excel); cells starting with = + - @ are prefixed so spreadsheets do not run them.</summary>
    public static byte[] Csv(ReportDocument report)
    {
        var sb = new StringBuilder();
        sb.AppendJoin(',', report.Columns.Select(c => Cell(c.Header))).Append("\r\n");
        foreach (var row in report.Rows)
            sb.AppendJoin(',', report.Columns.Select(c => Cell(ReportValues.Text(ReportValues.Get(row, c), c.Format)))).Append("\r\n");
        return [.. Encoding.UTF8.GetPreamble(), .. Encoding.UTF8.GetBytes(sb.ToString())];
    }

    private static string Cell(string text)
    {
        if (text.Length > 0 && "=+-@".Contains(text[0]) && !decimal.TryParse(text, out _)) text = "'" + text;
        return text.IndexOfAny([',', '"', '\r', '\n']) >= 0 ? $"\"{text.Replace("\"", "\"\"")}\"" : text;
    }
}
