using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;

namespace Cis.Api.Documents;

/// <summary>
/// Writes a <see cref="ReportDocument"/> as a PDF table (replaces the Crystal Reports and DataReport print-outs of listings):
/// title, subtitles, repeated header row, right-aligned numbers, totals row, page numbers.
/// </summary>
public sealed class PdfReportWriter
{
    /// <summary>MIME type of PDF.</summary>
    public const string ContentType = "application/pdf";

    /// <summary>Font used for the documents.</summary>
    public const string FontName = "Arial";

    /// <summary>The PDF bytes.</summary>
    public byte[] Write(ReportDocument report)
    {
        var document = new Document();
        document.Info.Title = report.Title;
        document.Styles[StyleNames.Normal]!.Font.Name = FontName;
        document.Styles[StyleNames.Normal]!.Font.Size = 8;

        var section = document.AddSection();
        section.PageSetup = document.DefaultPageSetup.Clone();
        section.PageSetup.PageFormat = PageFormat.A4;
        section.PageSetup.Orientation = report.Landscape ? Orientation.Landscape : Orientation.Portrait;
        section.PageSetup.LeftMargin = section.PageSetup.RightMargin = Unit.FromCentimeter(1.2);
        section.PageSetup.TopMargin = section.PageSetup.BottomMargin = Unit.FromCentimeter(1.5);

        var title = section.AddParagraph(report.Title);
        title.Format.Font.Size = 13;
        title.Format.Font.Bold = true;
        foreach (var line in report.Subtitles) section.AddParagraph(line).Format.Font.Color = Colors.DimGray;
        section.AddParagraph().Format.SpaceAfter = Unit.FromPoint(4);

        var footer = section.Footers.Primary.AddParagraph();
        footer.Format.Alignment = ParagraphAlignment.Right;
        footer.AddText("Page ");
        footer.AddPageField();
        footer.AddText(" of ");
        footer.AddNumPagesField();

        AddTable(section, report);

        var renderer = new PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();
        using var stream = new MemoryStream();
        renderer.PdfDocument.Save(stream, false);
        return stream.ToArray();
    }

    private static void AddTable(Section section, ReportDocument report)
    {
        var columns = report.Columns.Count == 0 ? [new ReportColumn("", "")] : report.Columns;
        var pageWidth = (report.Landscape ? 29.7 : 21.0) - 2.4;
        var table = section.AddTable();
        table.Borders.Width = 0.25;
        table.Borders.Color = Colors.LightGray;
        foreach (var _ in columns) table.AddColumn(Unit.FromCentimeter(pageWidth / columns.Count));

        var header = table.AddRow();
        header.HeadingFormat = true;
        header.Format.Font.Bold = true;
        header.Shading.Color = new Color(0xF1, 0xF5, 0xF7);
        for (var c = 0; c < columns.Count; c++) header.Cells[c].AddParagraph(columns[c].Header);

        foreach (var data in report.Rows)
        {
            var row = table.AddRow();
            for (var c = 0; c < columns.Count; c++) Fill(row.Cells[c], ReportValues.Text(ReportValues.Get(data, columns[c]), columns[c].Format), columns[c].Format);
        }

        if (!columns.Any(c => c.Total)) return;
        var total = table.AddRow();
        total.Format.Font.Bold = true;
        total.Cells[0].AddParagraph("Total");
        for (var c = 0; c < columns.Count; c++)
            if (columns[c].Total) Fill(total.Cells[c], ReportValues.Text(ReportValues.Sum(report, columns[c]), columns[c].Format), columns[c].Format);
    }

    private static void Fill(Cell cell, string text, ReportFormat format)
    {
        var paragraph = cell.AddParagraph(text);
        if (format is ReportFormat.Money or ReportFormat.Integer or ReportFormat.Number) paragraph.Format.Alignment = ParagraphAlignment.Right;
    }
}
