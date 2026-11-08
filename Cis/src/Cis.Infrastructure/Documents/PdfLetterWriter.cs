using Cis.Application.Documents;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;

namespace Cis.Infrastructure.Documents;

/// <summary>Writes a <see cref="LetterDocument"/> as an A4 PDF letter: heading, title, a two-column field table, paragraphs.</summary>
public sealed class PdfLetterWriter : ILetterWriter
{
    /// <inheritdoc />
    public byte[] Write(LetterDocument letter)
    {
        var document = new Document();
        document.Info.Title = letter.Title;
        document.Styles[StyleNames.Normal]!.Font.Name = PdfReportWriter.FontName;
        document.Styles[StyleNames.Normal]!.Font.Size = 10;
        var section = document.AddSection();
        section.PageSetup = document.DefaultPageSetup.Clone();
        section.PageSetup.PageFormat = PageFormat.A4;
        section.PageSetup.LeftMargin = section.PageSetup.RightMargin = Unit.FromCentimeter(2);
        section.PageSetup.TopMargin = section.PageSetup.BottomMargin = Unit.FromCentimeter(2);

        foreach (var line in letter.Heading) section.AddParagraph(line);
        var title = section.AddParagraph(letter.Title);
        title.Format.Font.Size = 13;
        title.Format.Font.Bold = true;
        title.Format.SpaceBefore = Unit.FromPoint(12);
        title.Format.SpaceAfter = Unit.FromPoint(10);

        var table = section.AddTable();
        table.AddColumn(Unit.FromCentimeter(5));
        table.AddColumn(Unit.FromCentimeter(12));
        foreach (var (label, value) in letter.Fields)
        {
            var row = table.AddRow();
            row.TopPadding = row.BottomPadding = Unit.FromPoint(2);
            row.Cells[0].AddParagraph(label).Format.Font.Bold = true;
            foreach (var part in (value ?? "").Replace("\r\n", "\n").Split('\n')) row.Cells[1].AddParagraph(part);
        }
        foreach (var p in letter.Paragraphs)
            section.AddParagraph(p).Format.SpaceBefore = Unit.FromPoint(8);

        var renderer = new PdfDocumentRenderer { Document = document };
        renderer.RenderDocument();
        using var stream = new MemoryStream();
        renderer.PdfDocument.Save(stream, false);
        return stream.ToArray();
    }
}
