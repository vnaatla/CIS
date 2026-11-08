namespace Cis.Application.Documents;

/// <summary>Writes reports in the format the user asked for (Excel / PDF / CSV).</summary>
public interface IReportExporter
{
    /// <summary>The report as a downloadable file.</summary>
    FileOutput Export(ReportDocument report, ReportFileType type);
}

/// <summary>A letter: heading lines, the filled-in fields (the Word template's bookmarks) and closing paragraphs.</summary>
/// <param name="Title">Title line.</param>
/// <param name="Heading">Lines above the title (letter head, date, reference).</param>
/// <param name="Fields">Label / value pairs in order.</param>
/// <param name="Paragraphs">Text after the fields.</param>
public record LetterDocument(string Title, IReadOnlyList<string> Heading, IReadOnlyList<(string Label, string Value)> Fields, IReadOnlyList<string> Paragraphs);

/// <summary>Writes letters as PDF (replaces the desktop's Word templates, which are not available).</summary>
public interface ILetterWriter
{
    /// <summary>The PDF bytes.</summary>
    byte[] Write(LetterDocument letter);
}

/// <summary>Finds copies of the desktop's legacy Excel templates (<c>\\svr-doc\HISREPORTS</c>).</summary>
public interface IReportTemplateStore
{
    /// <summary>The full path of the template, or null when it is not available.</summary>
    string? Find(string template);
}

/// <summary>Reads uploaded spreadsheets (the desktop opened them through Excel automation).</summary>
public interface ISpreadsheetReader
{
    /// <summary>The first sheet's rows from row 2 while column A is not empty (at most <paramref name="maxRows"/>); cells as text, dates as dd/MM/yyyy.
    /// Item 1 of each row is its sheet row number; columns are 1-based (index 0 unused).</summary>
    List<(int Row, string[] Cells)> Read(Stream content, int maxColumns, int maxRows = 60000);
}
