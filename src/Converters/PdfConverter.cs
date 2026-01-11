using System.Runtime.InteropServices;
using PuppeteerSharp;

namespace msguru.Converters;

/// <summary>
/// Converts documents to PDF format using PuppeteerSharp with the system's Chrome browser.
/// </summary>
public static class PdfConverter
{
    /// <summary>
    /// Converts HTML to PDF using Chrome/Chromium browser.
    /// </summary>
    /// <param name="htmlPath">The path to the input HTML file.</param>
    /// <param name="pdfPath">The path where the PDF file will be saved.</param>
    /// <param name="chromeExecutablePath">Optional path to Chrome executable. If null, attempts to find Chrome automatically.</param>
    /// <returns>A ConversionResult indicating success or failure.</returns>
    public static async Task<ConversionResult> HtmlToPdfAsync(
        string htmlPath,
        string pdfPath,
        string? chromeExecutablePath = null
    )
    {
        try
        {
            if (!File.Exists(htmlPath))
                return ConversionResult.FailureResult($"Input file not found: {htmlPath}");

            // Find Chrome executable if not provided
            var chromePath = chromeExecutablePath ?? FindChromeExecutable();
            if (string.IsNullOrEmpty(chromePath))
            {
                return ConversionResult.FailureResult(
                    "Chrome/Chromium not found. Please install Chrome or provide the executable path."
                );
            }

            if (!File.Exists(chromePath))
            {
                return ConversionResult.FailureResult(
                    $"Chrome executable not found at: {chromePath}"
                );
            }

            // Launch browser using system Chrome
            await using var browser = await Puppeteer.LaunchAsync(
                new LaunchOptions
                {
                    Headless = true,
                    ExecutablePath = chromePath,
                    Args = new[] { "--no-sandbox", "--disable-setuid-sandbox" },
                }
            );

            await using var page = await browser.NewPageAsync();

            // Load HTML file
            var htmlContent = await File.ReadAllTextAsync(htmlPath);
            await page.SetContentAsync(htmlContent);

            // Generate PDF
            await page.PdfAsync(
                pdfPath,
                new PdfOptions
                {
                    Format = PuppeteerSharp.Media.PaperFormat.A4,
                    PrintBackground = true,
                    MarginOptions = new PuppeteerSharp.Media.MarginOptions
                    {
                        Top = "1cm",
                        Right = "1cm",
                        Bottom = "1cm",
                        Left = "1cm",
                    },
                }
            );

            return ConversionResult.SuccessResult(pdfPath);
        }
        catch (Exception ex)
        {
            return ConversionResult.FailureResult($"PDF generation failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Converts DOCX to PDF by first converting to HTML, then using Chrome to render PDF.
    /// </summary>
    /// <param name="docxPath">The path to the input DOCX file.</param>
    /// <param name="pdfPath">The path where the PDF file will be saved.</param>
    /// <param name="chromeExecutablePath">Optional path to Chrome executable.</param>
    /// <returns>A ConversionResult indicating success or failure.</returns>
    public static async Task<ConversionResult> DocxToPdfAsync(
        string docxPath,
        string pdfPath,
        string? chromeExecutablePath = null
    )
    {
        try
        {
            if (!File.Exists(docxPath))
                return ConversionResult.FailureResult($"Input file not found: {docxPath}");

            // Create temporary HTML file
            var tempHtml = Path.GetTempFileName() + ".html";

            try
            {
                // Convert DOCX to HTML
                var htmlResult = DocxConverter.ToHtml(docxPath, tempHtml);
                if (!htmlResult.Success)
                    return htmlResult;

                // Convert HTML to PDF using Chrome
                var pdfResult = await HtmlToPdfAsync(tempHtml, pdfPath, chromeExecutablePath);
                return pdfResult;
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempHtml))
                    File.Delete(tempHtml);
            }
        }
        catch (Exception ex)
        {
            return ConversionResult.FailureResult($"PDF generation failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Converts XLSX to PDF by first converting to HTML, then using Chrome to render PDF.
    /// </summary>
    /// <param name="xlsxPath">The path to the input XLSX file.</param>
    /// <param name="pdfPath">The path where the PDF file will be saved.</param>
    /// <param name="chromeExecutablePath">Optional path to Chrome executable.</param>
    /// <returns>A ConversionResult indicating success or failure.</returns>
    public static async Task<ConversionResult> XlsxToPdfAsync(
        string xlsxPath,
        string pdfPath,
        string? chromeExecutablePath = null
    )
    {
        try
        {
            if (!File.Exists(xlsxPath))
                return ConversionResult.FailureResult($"Input file not found: {xlsxPath}");

            // Create temporary HTML file
            var tempHtml = Path.GetTempFileName() + ".html";

            try
            {
                // Convert XLSX to HTML
                var htmlResult = XlsxConverter.ToHtml(xlsxPath, tempHtml);
                if (!htmlResult.Success)
                    return htmlResult;

                // Convert HTML to PDF using Chrome
                var pdfResult = await HtmlToPdfAsync(tempHtml, pdfPath, chromeExecutablePath);
                return pdfResult;
            }
            finally
            {
                // Clean up temporary file
                if (File.Exists(tempHtml))
                    File.Delete(tempHtml);
            }
        }
        catch (Exception ex)
        {
            return ConversionResult.FailureResult($"PDF generation failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Attempts to find the Chrome/Chromium executable on the system.
    /// </summary>
    /// <returns>The path to Chrome executable, or null if not found.</returns>
    private static string? FindChromeExecutable()
    {
        var possiblePaths = new List<string>();

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            // Windows paths
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var programFilesX86 = Environment.GetFolderPath(
                Environment.SpecialFolder.ProgramFilesX86
            );
            var localAppData = Environment.GetFolderPath(
                Environment.SpecialFolder.LocalApplicationData
            );

            possiblePaths.AddRange(
                new[]
                {
                    Path.Combine(programFiles, @"Google\Chrome\Application\chrome.exe"),
                    Path.Combine(programFilesX86, @"Google\Chrome\Application\chrome.exe"),
                    Path.Combine(localAppData, @"Google\Chrome\Application\chrome.exe"),
                    Path.Combine(programFiles, @"Microsoft\Edge\Application\msedge.exe"),
                    Path.Combine(programFilesX86, @"Microsoft\Edge\Application\msedge.exe"),
                }
            );
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            // Linux paths
            possiblePaths.AddRange(
                new[]
                {
                    "/usr/bin/google-chrome",
                    "/usr/bin/google-chrome-stable",
                    "/usr/bin/chromium",
                    "/usr/bin/chromium-browser",
                    "/snap/bin/chromium",
                    "/usr/bin/microsoft-edge",
                    "/usr/bin/microsoft-edge-stable",
                }
            );
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            // macOS paths
            possiblePaths.AddRange(
                new[]
                {
                    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                    "/Applications/Chromium.app/Contents/MacOS/Chromium",
                    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
                }
            );
        }

        // Return first existing path
        return possiblePaths.FirstOrDefault(File.Exists);
    }
}
