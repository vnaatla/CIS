using System.Globalization;
using Cis.Application.Documents;
using ClosedXML.Excel;

namespace Cis.Infrastructure.Documents;

/// <summary><see cref="ISpreadsheetReader"/> with ClosedXML.</summary>
public sealed class SpreadsheetReader : ISpreadsheetReader
{
    /// <inheritdoc />
    public List<(int Row, string[] Cells)> Read(Stream content, int maxColumns, int maxRows = 60000, int keyColumn = 1, bool skipBlankKey = false)
    {
        using var book = new XLWorkbook(content);
        var sheet = book.Worksheet(1);
        var rows = new List<(int, string[])>();
        var last = Math.Min(maxRows - 1, sheet.LastRowUsed()?.RowNumber() ?? 1);
        for (var r = 2; r < maxRows; r++)
        {
            if (skipBlankKey && r > last) break;
            if (Text(sheet.Cell(r, keyColumn)).Length == 0)
            {
                if (skipBlankKey) continue;
                break;
            }
            var cells = new string[maxColumns + 1];
            for (var c = 1; c <= maxColumns; c++) cells[c] = Text(sheet.Cell(r, c));
            rows.Add((r, cells));
        }
        return rows;
    }

    private static string Text(IXLCell cell) => cell.DataType switch
    {
        XLDataType.DateTime => cell.GetDateTime().ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
        XLDataType.Number => cell.GetDouble().ToString(CultureInfo.InvariantCulture),
        _ => cell.GetFormattedString().Trim(),
    };
}
