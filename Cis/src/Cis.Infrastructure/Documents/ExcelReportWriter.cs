using Cis.Application.Documents;
using ClosedXML.Excel;

namespace Cis.Infrastructure.Documents;

/// <summary>
/// Writes a <see cref="ReportDocument"/> as .xlsx (replaces ModListviewtoExcel.ExportToExcel* and the .xlt templates). With a template from
/// <see cref="IReportTemplateStore"/> the rows are written into the template's first sheet from <see cref="ReportDocument.TemplateStartRow"/>;
/// otherwise a sheet is generated: title, subtitles, bold header, data, optional totals row, frozen header, auto-fitted columns.
/// </summary>
public sealed class ExcelReportWriter(IReportTemplateStore templates)
{
    /// <summary>MIME type of .xlsx.</summary>
    public const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /// <summary>The workbook bytes.</summary>
    public byte[] Write(ReportDocument report)
    {
        var template = report.Template is { } name ? templates.Find(name) : null;
        using var book = template is null ? new XLWorkbook() : new XLWorkbook(template);
        WriteSheet(template is null ? book.Worksheets.Add(SheetName(report.Title)) : book.Worksheet(1), report, template is null);
        foreach (var extra in report.ExtraSheets) WriteSheet(book.Worksheets.Add(SheetName(extra.Title)), extra, true);

        using var stream = new MemoryStream();
        book.SaveAs(stream);
        return stream.ToArray();
    }

    private static void WriteSheet(IXLWorksheet sheet, ReportDocument report, bool heading)
    {
        var row = heading ? WriteHeading(sheet, report) : report.TemplateStartRow;
        var firstData = row;
        foreach (var data in report.Rows)
        {
            for (var c = 0; c < report.Columns.Count; c++) SetCell(sheet.Cell(row, c + 1), ReportValues.Get(data, report.Columns[c]), report.Columns[c].Format);
            row++;
        }

        if (report.Columns.Any(c => c.Total))
        {
            sheet.Cell(row, 1).Value = "Total";
            for (var c = 0; c < report.Columns.Count; c++)
            {
                if (!report.Columns[c].Total) continue;
                var cell = sheet.Cell(row, c + 1);
                cell.Value = ReportValues.Sum(report, report.Columns[c]);
                cell.Style.NumberFormat.Format = report.Columns[c].Format == ReportFormat.Money ? "#,##0.00" : "0";
            }
            sheet.Row(row).Style.Font.Bold = true;
        }

        if (heading)
        {
            sheet.SheetView.FreezeRows(firstData - 1);
            sheet.Columns(1, Math.Max(1, report.Columns.Count)).AdjustToContents(1, Math.Min(row, firstData + 500));
        }
    }

    private static int WriteHeading(IXLWorksheet sheet, ReportDocument report)
    {
        var row = 1;
        sheet.Cell(row, 1).Value = report.Title;
        sheet.Cell(row, 1).Style.Font.Bold = true;
        sheet.Cell(row, 1).Style.Font.FontSize = 14;
        row++;
        foreach (var line in report.Subtitles) sheet.Cell(row++, 1).Value = line;
        row++;
        for (var c = 0; c < report.Columns.Count; c++) sheet.Cell(row, c + 1).Value = report.Columns[c].Header;
        var header = sheet.Range(row, 1, row, Math.Max(1, report.Columns.Count));
        header.Style.Font.Bold = true;
        header.Style.Fill.BackgroundColor = XLColor.FromHtml("#F1F5F7");
        header.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        return row + 1;
    }

    private static void SetCell(IXLCell cell, object? value, ReportFormat format)
    {
        switch (format)
        {
            case ReportFormat.Money or ReportFormat.Integer or ReportFormat.Number when ReportValues.Number(value) is { } n:
                cell.Value = n;
                cell.Style.NumberFormat.Format = format switch { ReportFormat.Money => "#,##0.00", ReportFormat.Integer => "0", _ => "General" };
                break;
            case ReportFormat.Date when ReportValues.Date(value) is { } d:
                cell.Value = d;
                cell.Style.DateFormat.Format = "dd/mm/yyyy";
                break;
            default:
                cell.Value = ReportValues.Text(value, format);
                break;
        }
    }

    private static string SheetName(string title)
    {
        var clean = new string(title.Where(ch => !"[]:*?/\\".Contains(ch)).ToArray()).Trim();
        return clean.Length == 0 ? "Report" : clean.Length > 31 ? clean[..31] : clean;
    }
}
