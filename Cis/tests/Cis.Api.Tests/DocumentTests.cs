using System.Text;
using Cis.Application.Documents;
using Cis.Application.Email;
using Cis.Application.Files;
using Cis.Domain.Errors;
using Cis.Infrastructure.Documents;
using Cis.Infrastructure.Email;
using Cis.Infrastructure.Files;
using ClosedXML.Excel;
using Microsoft.Extensions.Options;

namespace Cis.Api.Tests;

public class DocumentTests
{
    private static readonly ReportColumn[] Columns =
    [
        new("ClaimNo", "Claim No"),
        new("Visit", "Visit Date", ReportFormat.Date),
        new("Amount", "Amount (RM)", ReportFormat.Money, Total: true),
    ];

    private static ReportDocument Report(string? template = null) => new("Claim Report", Columns,
    [
        new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) { ["claimno"] = "C001", ["Visit"] = new DateTime(2026, 10, 1), ["Amount"] = 12.5m },
        new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase) { ["ClaimNo"] = "=HYPERLINK(\"x\")", ["Visit"] = null, ["Amount"] = "1000" },
    ])
    {
        Subtitles = ["Payor: ALL"],
        FileName = "ClaimReport",
        Template = template,
    };

    private static IReportExporter Exporter(string? templateRoot = null) =>
        new ReportExporter(new ExcelReportWriter(new ReportTemplateStore(Options.Create(new ReportOptions { TemplateRoot = templateRoot }))), new PdfReportWriter());

    [Fact]
    public void Excel_has_title_header_rows_and_a_totals_row()
    {
        var file = Exporter().Export(Report(), ReportFileType.Excel);
        Assert.Equal("ClaimReport.xlsx", file.FileName);
        using var book = new XLWorkbook(new MemoryStream(file.Content));
        var sheet = book.Worksheet(1);
        Assert.Equal("Claim Report", sheet.Cell(1, 1).GetString());
        Assert.Equal("Payor: ALL", sheet.Cell(2, 1).GetString());
        Assert.Equal("Amount (RM)", sheet.Cell(4, 3).GetString());
        Assert.Equal("C001", sheet.Cell(5, 1).GetString());
        Assert.Equal(new DateTime(2026, 10, 1), sheet.Cell(5, 2).GetDateTime());
        Assert.Equal(1012.5, sheet.Cell(7, 3).GetDouble());
        Assert.Equal("Total", sheet.Cell(7, 1).GetString());
    }

    [Fact]
    public void Excel_fills_a_legacy_template_when_one_is_available()
    {
        var root = Path.Combine(Path.GetTempPath(), $"cis-tpl-{Guid.NewGuid():N}");
        Directory.CreateDirectory(Path.Combine(root, "CISTemplate"));
        using (var template = new XLWorkbook())
        {
            template.AddWorksheet("Claims").Cell(1, 1).Value = "TEMPLATE HEADER";
            template.SaveAs(Path.Combine(root, "CISTemplate", "ClaimReport.xlsx"));
        }

        var file = Exporter(root).Export(Report(@"CISTemplate\ClaimReport.xlt"), ReportFileType.Excel);
        using var book = new XLWorkbook(new MemoryStream(file.Content));
        Assert.Equal("TEMPLATE HEADER", book.Worksheet(1).Cell(1, 1).GetString());
        Assert.Equal("C001", book.Worksheet(1).Cell(2, 1).GetString());
        Directory.Delete(root, true);
    }

    [Fact]
    public void Template_store_refuses_paths_outside_the_root()
    {
        var store = new ReportTemplateStore(Options.Create(new ReportOptions { TemplateRoot = Path.GetTempPath() }));
        Assert.Null(store.Find(@"..\..\Windows\win"));
        Assert.Null(new ReportTemplateStore(Options.Create(new ReportOptions())).Find("CISTemplate\\x"));
    }

    [Fact]
    public void Csv_quotes_cells_and_neutralises_formulas()
    {
        var text = Encoding.UTF8.GetString(Exporter().Export(Report(), ReportFileType.Csv).Content);
        var lines = text.TrimStart('\uFEFF').Split("\r\n");
        Assert.Equal("Claim No,Visit Date,Amount (RM)", lines[0]);
        Assert.Equal("C001,01/10/2026,12.50", lines[1]);
        Assert.Equal("\"'=HYPERLINK(\"\"x\"\")\",,\"1,000.00\"", lines[2]);
    }

    [Fact]
    public void Pdf_is_a_pdf_document()
    {
        PdfSharp.Fonts.GlobalFontSettings.FontResolver ??= new PdfFontResolver();
        var file = Exporter().Export(Report(), ReportFileType.Pdf);
        Assert.Equal("application/pdf", file.ContentType);
        Assert.Equal("%PDF", Encoding.ASCII.GetString(file.Content, 0, 4));
    }

    [Fact]
    public async Task File_drop_writes_only_to_configured_folders_and_rejects_paths()
    {
        var root = Path.Combine(Path.GetTempPath(), $"cis-drop-{Guid.NewGuid():N}");
        var drop = new FileDrop(Options.Create(new FileDropOptions { Folders = new(StringComparer.OrdinalIgnoreCase) { [FileFolders.ClaimAscii] = root } }));

        Assert.Equal(Path.Combine(root, "a.txt"), await drop.SaveAsync(FileFolders.ClaimAscii, "a.txt", [1, 2]));
        Assert.Null(await drop.SaveAsync(FileFolders.Bordx, "a.txt", [1]));
        await Assert.ThrowsAsync<ArgumentException>(() => drop.SaveAsync(FileFolders.ClaimAscii, "..", [1]));
        Directory.Delete(root, true);
    }

    [Fact]
    public async Task Email_goes_to_the_pickup_folder_or_reports_that_it_is_not_configured()
    {
        var pickup = Path.Combine(Path.GetTempPath(), $"cis-mail-{Guid.NewGuid():N}");
        var message = new EmailMessage(["a@b.com"], "Subject", "<p>Body</p>") { Attachments = [("x.txt", [65])] };

        await new SmtpEmailSender(Options.Create(new EmailOptions { PickupDirectory = pickup })).SendAsync(message);
        Assert.Contains("Subject: Subject", await File.ReadAllTextAsync(Assert.Single(Directory.GetFiles(pickup, "*.eml"))));
        Directory.Delete(pickup, true);

        await Assert.ThrowsAsync<NotConfiguredException>(() => new SmtpEmailSender(Options.Create(new EmailOptions())).SendAsync(message));
    }
}
