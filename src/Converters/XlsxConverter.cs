using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;

namespace msguru.Converters;

/// <summary>
/// Converts XLSX spreadsheets to and from HTML format.
/// </summary>
public static class XlsxConverter
{
    /// <summary>
    /// Converts an XLSX spreadsheet to HTML.
    /// </summary>
    /// <param name="xlsxPath">The path to the input XLSX file.</param>
    /// <param name="htmlPath">The path where the HTML file will be saved.</param>
    /// <returns>A ConversionResult indicating success or failure.</returns>
    public static ConversionResult ToHtml(string xlsxPath, string htmlPath)
    {
        try
        {
            if (!File.Exists(xlsxPath))
                return ConversionResult.FailureResult($"Input file not found: {xlsxPath}");

            using var doc = SpreadsheetDocument.Open(xlsxPath, false);
            var workbookPart = doc.WorkbookPart;

            if (workbookPart == null)
                return ConversionResult.FailureResult("Workbook part is missing or invalid.");

            var html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("    <meta charset=\"utf-8\">");
            html.AppendLine("    <title>Spreadsheet</title>");
            html.AppendLine("    <style>");
            html.AppendLine(
                "        body { font-family: Calibri, Arial, sans-serif; margin: 40px; }"
            );
            html.AppendLine("        table { border-collapse: collapse; margin: 20px 0; }");
            html.AppendLine(
                "        td, th { border: 1px solid #ccc; padding: 8px; text-align: left; }"
            );
            html.AppendLine("        th { background-color: #f0f0f0; font-weight: bold; }");
            html.AppendLine("        h2 { margin: 30px 0 10px 0; }");
            html.AppendLine("    </style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");

            int totalSheets = 0;
            var sheets = workbookPart.Workbook.Descendants<Sheet>();

            foreach (var sheet in sheets)
            {
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                if (sheetData == null)
                    continue;

                html.AppendLine(
                    $"    <h2>{System.Net.WebUtility.HtmlEncode(sheet.Name?.Value ?? "Sheet")}</h2>"
                );
                html.AppendLine("    <table>");

                var rows = sheetData.Elements<Row>().ToList();
                var firstRow = true;

                foreach (var row in rows)
                {
                    html.AppendLine("        <tr>");

                    foreach (var cell in row.Elements<Cell>())
                    {
                        var cellValue = GetCellValue(cell, workbookPart);
                        var tag = firstRow ? "th" : "td";
                        html.AppendLine(
                            $"            <{tag}>{System.Net.WebUtility.HtmlEncode(cellValue)}</{tag}>"
                        );
                    }

                    html.AppendLine("        </tr>");
                    firstRow = false;
                }

                html.AppendLine("    </table>");
                totalSheets++;
            }

            html.AppendLine("</body>");
            html.AppendLine("</html>");

            File.WriteAllText(htmlPath, html.ToString());

            return ConversionResult.SuccessResult(htmlPath, totalSheets);
        }
        catch (Exception ex)
        {
            return ConversionResult.FailureResult($"Conversion failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Converts HTML tables to an XLSX spreadsheet.
    /// </summary>
    /// <param name="htmlPath">The path to the input HTML file.</param>
    /// <param name="xlsxPath">The path where the XLSX file will be saved.</param>
    /// <returns>A ConversionResult indicating success or failure.</returns>
    public static ConversionResult FromHtml(string htmlPath, string xlsxPath)
    {
        try
        {
            if (!File.Exists(htmlPath))
                return ConversionResult.FailureResult($"Input file not found: {htmlPath}");

            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(htmlPath);

            using var spreadsheet = SpreadsheetDocument.Create(
                xlsxPath,
                DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook
            );
            var workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var tables = htmlDoc.DocumentNode.SelectNodes("//table");

            if (tables == null || tables.Count == 0)
                return ConversionResult.FailureResult("No tables found in HTML document.");

            uint sheetId = 1;

            foreach (var table in tables)
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                var sheet = new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId,
                    Name = $"Sheet{sheetId}",
                };
                sheets.Append(sheet);

                var htmlRows = table.SelectNodes(".//tr");
                if (htmlRows != null)
                {
                    uint rowIndex = 1;
                    foreach (var htmlRow in htmlRows)
                    {
                        var row = new Row { RowIndex = rowIndex };

                        var cells = htmlRow.SelectNodes(".//td|.//th");
                        if (cells != null)
                        {
                            foreach (var htmlCell in cells)
                            {
                                var cellText = System
                                    .Net.WebUtility.HtmlDecode(htmlCell.InnerText)
                                    .Trim();
                                var cell = new Cell
                                {
                                    CellValue = new CellValue(cellText),
                                    DataType = CellValues.String,
                                };
                                row.Append(cell);
                            }
                        }

                        sheetData.Append(row);
                        rowIndex++;
                    }
                }

                sheetId++;
            }

            workbookPart.Workbook.Save();

            return ConversionResult.SuccessResult(xlsxPath, tables.Count);
        }
        catch (Exception ex)
        {
            return ConversionResult.FailureResult($"Conversion failed: {ex.Message}");
        }
    }

    private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.CellValue == null)
            return string.Empty;

        var value = cell.CellValue.Text;

        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (stringTable != null)
            {
                return stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }
        }

        return value ?? string.Empty;
    }
}
