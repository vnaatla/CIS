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
