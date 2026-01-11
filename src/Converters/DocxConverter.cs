using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;

namespace msguru.Converters;

/// <summary>
/// Converts DOCX documents to and from HTML format.
/// </summary>
public static class DocxConverter
{
    /// <summary>
    /// Converts a DOCX document to HTML.
    /// </summary>
    /// <param name="docxPath">The path to the input DOCX file.</param>
    /// <param name="htmlPath">The path where the HTML file will be saved.</param>
    /// <returns>A ConversionResult indicating success or failure.</returns>
    public static ConversionResult ToHtml(string docxPath, string htmlPath)
    {
        try
        {
            if (!File.Exists(docxPath))
                return ConversionResult.FailureResult($"Input file not found: {docxPath}");

            using var doc = WordprocessingDocument.Open(docxPath, false);
            var body = doc.MainDocumentPart?.Document.Body;

            if (body == null)
                return ConversionResult.FailureResult("Document body is empty or invalid.");

            var html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("    <meta charset=\"utf-8\">");
            html.AppendLine("    <title>Document</title>");
            html.AppendLine("    <style>");
            html.AppendLine(
                "        body { font-family: Calibri, Arial, sans-serif; margin: 40px; }"
            );
            html.AppendLine("        p { margin: 0 0 10px 0; }");
            html.AppendLine("        h1, h2, h3, h4, h5, h6 { margin: 20px 0 10px 0; }");
            html.AppendLine("        table { border-collapse: collapse; margin: 10px 0; }");
            html.AppendLine("        td, th { border: 1px solid #ccc; padding: 5px; }");
            html.AppendLine("    </style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");

            int paragraphCount = 0;

            foreach (var element in body.Elements())
            {
                if (element is Paragraph para)
                {
                    var text = para.InnerText;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        // Check for heading style
                        var styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                        var tag = styleId switch
                        {
                            "Heading1" => "h1",
                            "Heading2" => "h2",
                            "Heading3" => "h3",
                            "Heading4" => "h4",
                            "Heading5" => "h5",
                            "Heading6" => "h6",
                            _ => "p",
                        };

                        html.AppendLine(
                            $"    <{tag}>{System.Net.WebUtility.HtmlEncode(text)}</{tag}>"
                        );
                        paragraphCount++;
                    }
                }
                else if (element is Table table)
                {
                    html.AppendLine("    <table>");
                    foreach (var row in table.Elements<TableRow>())
                    {
                        html.AppendLine("        <tr>");
                        foreach (var cell in row.Elements<TableCell>())
                        {
                            var cellText = cell.InnerText;
                            html.AppendLine(
                                $"            <td>{System.Net.WebUtility.HtmlEncode(cellText)}</td>"
                            );
                        }
                        html.AppendLine("        </tr>");
                    }
                    html.AppendLine("    </table>");
                    paragraphCount++;
                }
            }

            html.AppendLine("</body>");
            html.AppendLine("</html>");

            File.WriteAllText(htmlPath, html.ToString());

            return ConversionResult.SuccessResult(htmlPath, paragraphCount);
        }
        catch (Exception ex)
        {
            return ConversionResult.FailureResult($"Conversion failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Converts HTML to a DOCX document.
    /// </summary>
    /// <param name="htmlPath">The path to the input HTML file.</param>
    /// <param name="docxPath">The path where the DOCX file will be saved.</param>
    /// <returns>A ConversionResult indicating success or failure.</returns>
    public static ConversionResult FromHtml(string htmlPath, string docxPath)
    {
        try
        {
            if (!File.Exists(htmlPath))
                return ConversionResult.FailureResult($"Input file not found: {htmlPath}");

            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(htmlPath);

            using var wordDoc = WordprocessingDocument.Create(
                docxPath,
                DocumentFormat.OpenXml.WordprocessingDocumentType.Document
            );
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
            var body = mainPart.Document.Body;

            if (body == null)
                return ConversionResult.FailureResult("Failed to create document body.");

            int elementCount = 0;

            // Parse HTML body elements
            var htmlBody = htmlDoc.DocumentNode.SelectSingleNode("//body");
            if (htmlBody == null)
                htmlBody = htmlDoc.DocumentNode; // Fallback to root if no body tag

            foreach (var node in htmlBody.ChildNodes)
            {
                if (node.NodeType != HtmlNodeType.Element)
                    continue;

                var text = System.Net.WebUtility.HtmlDecode(node.InnerText).Trim();
                if (string.IsNullOrWhiteSpace(text))
                    continue;

                var paragraph = new Paragraph();
                var run = new Run(new Text(text));

                // Apply heading styles
                if (node.Name.StartsWith("h") && node.Name.Length == 2)
                {
                    var level = node.Name[1] - '0';
                    if (level >= 1 && level <= 6)
                    {
                        var props = new ParagraphProperties(
                            new ParagraphStyleId { Val = $"Heading{level}" }
                        );
                        paragraph.ParagraphProperties = props;
                    }
                }

                paragraph.Append(run);
                body.Append(paragraph);
                elementCount++;
            }

            mainPart.Document.Save();

            return ConversionResult.SuccessResult(docxPath, elementCount);
        }
        catch (Exception ex)
        {
            return ConversionResult.FailureResult($"Conversion failed: {ex.Message}");
        }
    }
}
