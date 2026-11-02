namespace Cis.Application.Documents;

/// <summary>Writes reports in the format the user asked for (Excel / PDF / CSV).</summary>
public interface IReportExporter
{
    /// <summary>The report as a downloadable file.</summary>
    FileOutput Export(ReportDocument report, ReportFileType type);
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
