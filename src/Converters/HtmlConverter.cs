using HtmlAgilityPack;
using System.Text;

namespace msguru.Converters;

/// <summary>
/// Provides HTML parsing and manipulation utilities.
/// </summary>
public static class HtmlConverter
{
    /// <summary>
    /// Validates and cleans HTML content.
    /// </summary>
    /// <param name="inputPath">The path to the input HTML file.</param>
    /// <param name="outputPath">The path where the cleaned HTML file will be saved.</param>
    /// <returns>A ConversionResult indicating success or failure.</returns>
    public static ConversionResult CleanHtml(string inputPath, string outputPath)
    {
        try
        {
            if (!File.Exists(inputPath))
                return ConversionResult.FailureResult($"Input file not found: {inputPath}");

            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(inputPath);

            // Remove script tags for security
            var scriptNodes = htmlDoc.DocumentNode.SelectNodes("//script");
            if (scriptNodes != null)
            {
                foreach (var node in scriptNodes)
                {
                    node.Remove();
                }
            }

            // Remove inline event handlers
            RemoveEventHandlers(htmlDoc.DocumentNode);

            htmlDoc.Save(outputPath);

            return ConversionResult.SuccessResult(outputPath);
        }
        catch (Exception ex)
        {
            return ConversionResult.FailureResult($"HTML cleaning failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Extracts text content from HTML, removing all tags.
    /// </summary>
    /// <param name="htmlPath">The path to the input HTML file.</param>
    /// <param name="textPath">The path where the plain text file will be saved.</param>
    /// <returns>A ConversionResult indicating success or failure.</returns>
    public static ConversionResult ToPlainText(string htmlPath, string textPath)
    {
        try
        {
            if (!File.Exists(htmlPath))
                return ConversionResult.FailureResult($"Input file not found: {htmlPath}");

            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(htmlPath);

            var text = htmlDoc.DocumentNode.InnerText;
            text = System.Net.WebUtility.HtmlDecode(text);

            File.WriteAllText(textPath, text);

            return ConversionResult.SuccessResult(textPath);
        }
        catch (Exception ex)
        {
            return ConversionResult.FailureResult($"Text extraction failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Extracts all links from HTML document.
    /// </summary>
    /// <param name="htmlPath">The path to the input HTML file.</param>
    /// <returns>A list of all href values found in anchor tags.</returns>
    public static List<string> ExtractLinks(string htmlPath)
    {
        var links = new List<string>();

        try
        {
            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(htmlPath);

            var linkNodes = htmlDoc.DocumentNode.SelectNodes("//a[@href]");
            if (linkNodes != null)
            {
                foreach (var node in linkNodes)
                {
                    var href = node.GetAttributeValue("href", string.Empty);
                    if (!string.IsNullOrWhiteSpace(href))
                    {
                        links.Add(href);
                    }
                }
            }
        }
        catch
        {
            // Return empty list on error
        }

        return links;
    }

    /// <summary>
    /// Converts plain text to HTML with basic formatting.
    /// </summary>
    /// <param name="textPath">The path to the input text file.</param>
    /// <param name="htmlPath">The path where the HTML file will be saved.</param>
    /// <returns>A ConversionResult indicating success or failure.</returns>
    public static ConversionResult FromPlainText(string textPath, string htmlPath)
    {
        try
        {
            if (!File.Exists(textPath))
                return ConversionResult.FailureResult($"Input file not found: {textPath}");

            var text = File.ReadAllText(textPath);
            var paragraphs = text.Split(new[] { "\r\n\r\n", "\n\n" }, StringSplitOptions.RemoveEmptyEntries);

            var html = new StringBuilder();
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("    <meta charset=\"utf-8\">");
            html.AppendLine("    <title>Document</title>");
            html.AppendLine("    <style>");
            html.AppendLine("        body { font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }");
            html.AppendLine("        p { margin: 0 0 15px 0; }");
            html.AppendLine("    </style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");

            foreach (var para in paragraphs)
            {
                var encoded = System.Net.WebUtility.HtmlEncode(para.Trim());
                // Convert single line breaks to <br>
                encoded = encoded.Replace("\r\n", "<br>").Replace("\n", "<br>");
                html.AppendLine($"    <p>{encoded}</p>");
            }

            html.AppendLine("</body>");
            html.AppendLine("</html>");

            File.WriteAllText(htmlPath, html.ToString());

            return ConversionResult.SuccessResult(htmlPath, paragraphs.Length);
        }
        catch (Exception ex)
        {
            return ConversionResult.FailureResult($"HTML generation failed: {ex.Message}");
        }
    }

    private static void RemoveEventHandlers(HtmlNode node)
    {
        if (node.NodeType == HtmlNodeType.Element)
        {
            var attributesToRemove = node.Attributes
                .Where(attr => attr.Name.StartsWith("on", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var attr in attributesToRemove)
            {
                node.Attributes.Remove(attr);
            }
        }

        foreach (var child in node.ChildNodes)
        {
            RemoveEventHandlers(child);
        }
    }
}
