using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace msguru.Interop;

/// <summary>
/// Singleton service for Excel COM Interop operations.
/// Provides methods for workbook manipulation, data extraction, and format conversion.
/// </summary>
public class ExcelService : IOfficeService
{
    private Excel.Application? _excelApp;
    private bool _disposed;
    private readonly object _lock = new();

    public string ApplicationName => "Excel";

    public bool IsApplicationRunning
    {
        get
        {
            lock (_lock)
            {
                return _excelApp != null;
            }
        }
    }

    public void Initialize()
    {
        lock (_lock)
        {
            if (_excelApp == null)
            {
                _excelApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    ScreenUpdating = false
                };
            }
        }
    }

    public void Cleanup()
    {
        lock (_lock)
        {
            if (_excelApp != null)
            {
                try
                {
                    _excelApp.Quit();
                }
                catch { }
                finally
                {
                    ServiceManager.ReleaseComObject(_excelApp);
                    _excelApp = null;
                }
            }
        }
    }

    /// <summary>
    /// Gets metadata and information about a workbook.
    /// </summary>
    /// <param name="path">The path to the Excel workbook file.</param>
    /// <returns>A WorkbookInfo object containing metadata and sheet information.</returns>
    /// <exception cref="FileNotFoundException">Thrown when the workbook file is not found.</exception>
    public WorkbookInfo GetWorkbookInfo(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"Workbook not found: {path}");

        Initialize();

        Excel.Workbook? workbook = null;
        try
        {
            workbook = _excelApp!.Workbooks.Open(path, ReadOnly: true);

            dynamic props = workbook.BuiltinDocumentProperties;
            var info = new WorkbookInfo
            {
                Path = path,
                FileName = Path.GetFileName(path),
                SheetCount = workbook.Worksheets.Count,
                SheetNames = new List<string>(),
                Author = workbook.Author,
                Title = workbook.Title,
                Subject = workbook.Subject,
                LastSavedBy = props["Last Author"].Value?.ToString() ?? string.Empty,
                Created = props["Creation Date"].Value is DateTime created ? created : DateTime.MinValue,
                Modified = props["Last Save Time"].Value is DateTime modified ? modified : DateTime.MinValue
            };

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                info.SheetNames.Add(sheet.Name);
                ServiceManager.ReleaseComObject(sheet);
            }

            return info;
        }
        finally
        {
            if (workbook != null)
            {
                workbook.Close(false);
                ServiceManager.ReleaseComObject(workbook);
            }
        }
    }

    /// <summary>
    /// Lists all worksheet names in a workbook.
    /// </summary>
    /// <param name="path">The path to the Excel workbook file.</param>
    /// <returns>A list of worksheet names.</returns>
    /// <exception cref="FileNotFoundException">Thrown when the workbook file is not found.</exception>
    public List<string> ListSheets(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"Workbook not found: {path}");

        Initialize();

        Excel.Workbook? workbook = null;
        try
        {
            workbook = _excelApp!.Workbooks.Open(path, ReadOnly: true);
            var sheetNames = new List<string>();

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                sheetNames.Add(sheet.Name);
                ServiceManager.ReleaseComObject(sheet);
            }

            return sheetNames;
        }
        finally
        {
            if (workbook != null)
            {
                workbook.Close(false);
                ServiceManager.ReleaseComObject(workbook);
            }
        }
    }

    /// <summary>
    /// Converts a workbook or specific sheet to CSV format.
    /// </summary>
    /// <param name="inputPath">The path to the Excel workbook file.</param>
    /// <param name="outputPath">The path where the CSV file will be saved.</param>
    /// <param name="sheetName">Optional name of the sheet to convert. If null, converts the first sheet.</param>
    /// <exception cref="FileNotFoundException">Thrown when the workbook file is not found.</exception>
    public void ConvertToCsv(string inputPath, string outputPath, string? sheetName = null)
    {
        if (!File.Exists(inputPath))
            throw new FileNotFoundException($"Workbook not found: {inputPath}");

        Initialize();

        Excel.Workbook? workbook = null;
        Excel.Worksheet? targetSheet = null;

        try
        {
            workbook = _excelApp!.Workbooks.Open(inputPath, ReadOnly: true);

            if (!string.IsNullOrEmpty(sheetName))
            {
                targetSheet = (Excel.Worksheet)workbook.Worksheets[sheetName];
            }
            else
            {
                targetSheet = (Excel.Worksheet)workbook.Worksheets[1];
            }

            // Save as CSV
            targetSheet.SaveAs(outputPath, Excel.XlFileFormat.xlCSV);
        }
        finally
        {
            if (targetSheet != null)
                ServiceManager.ReleaseComObject(targetSheet);

            if (workbook != null)
            {
                workbook.Close(false);
                ServiceManager.ReleaseComObject(workbook);
            }
        }
    }

    /// <summary>
    /// Extracts a specific sheet to a new workbook file.
    /// </summary>
    /// <param name="inputPath">The path to the source Excel workbook file.</param>
    /// <param name="sheetName">The name of the sheet to extract.</param>
    /// <param name="outputPath">The path where the new workbook will be saved.</param>
    /// <exception cref="FileNotFoundException">Thrown when the workbook file is not found.</exception>
    public void ExtractSheet(string inputPath, string sheetName, string outputPath)
    {
        if (!File.Exists(inputPath))
            throw new FileNotFoundException($"Workbook not found: {inputPath}");

        Initialize();

        Excel.Workbook? sourceWorkbook = null;
        Excel.Workbook? targetWorkbook = null;
        Excel.Worksheet? sheet = null;

        try
        {
            sourceWorkbook = _excelApp!.Workbooks.Open(inputPath, ReadOnly: true);
            sheet = (Excel.Worksheet)sourceWorkbook.Worksheets[sheetName];

            // Create new workbook
            targetWorkbook = _excelApp.Workbooks.Add();

            // Copy sheet
            sheet.Copy(targetWorkbook.Worksheets[1]);

            // Save
            targetWorkbook.SaveAs(outputPath);
        }
        finally
        {
            if (sheet != null)
                ServiceManager.ReleaseComObject(sheet);

            if (targetWorkbook != null)
            {
                targetWorkbook.Close(true);
                ServiceManager.ReleaseComObject(targetWorkbook);
            }

            if (sourceWorkbook != null)
            {
                sourceWorkbook.Close(false);
                ServiceManager.ReleaseComObject(sourceWorkbook);
            }
        }
    }

    /// <summary>
    /// Refreshes all data connections in a workbook.
    /// </summary>
    /// <summary>
    /// Refreshes all data connections in a workbook.
    /// </summary>
    /// <param name="path">The path to the Excel workbook file.</param>
    /// <exception cref="FileNotFoundException">Thrown when the workbook file is not found.</exception>
    public void RefreshConnections(string path)
    {
        if (!File.Exists(path))
            throw new FileNotFoundException($"Workbook not found: {path}");

        Initialize();

        Excel.Workbook? workbook = null;
        try
        {
            workbook = _excelApp!.Workbooks.Open(path);
            workbook.RefreshAll();
            workbook.Save();
        }
        finally
        {
            if (workbook != null)
            {
                workbook.Close(true);
                ServiceManager.ReleaseComObject(workbook);
            }
        }
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            Cleanup();
            _disposed = true;
        }
        GC.SuppressFinalize(this);
    }

    ~ExcelService()
    {
        Dispose();
    }
}

/// <summary>
/// Contains metadata and information about an Excel workbook.
/// </summary>
public class WorkbookInfo
{
    /// <summary>Gets or sets the full path to the workbook file.</summary>
    public string Path { get; set; } = string.Empty;

    /// <summary>Gets or sets the file name of the workbook.</summary>
    public string FileName { get; set; } = string.Empty;

    /// <summary>Gets or sets the number of worksheets in the workbook.</summary>
    public int SheetCount { get; set; }

    /// <summary>Gets or sets the list of worksheet names.</summary>
    public List<string> SheetNames { get; set; } = new();

    /// <summary>Gets or sets the author of the workbook.</summary>
    public string Author { get; set; } = string.Empty;

    /// <summary>Gets or sets the title of the workbook.</summary>
    public string Title { get; set; } = string.Empty;

    /// <summary>Gets or sets the subject of the workbook.</summary>
    public string Subject { get; set; } = string.Empty;

    /// <summary>Gets or sets the name of the user who last saved the workbook.</summary>
    public string LastSavedBy { get; set; } = string.Empty;

    /// <summary>Gets or sets the creation date of the workbook.</summary>
    public DateTime Created { get; set; }

    /// <summary>Gets or sets the last modified date of the workbook.</summary>
    public DateTime Modified { get; set; }
}
