using System.CommandLine;
using msguru.Interop;

namespace msguru.Commands;

/// <summary>
/// Provides commands for Excel workbook operations using COM Interop.
/// </summary>
public static class ExcelCommand
{
    /// <summary>
    /// Gets the main Excel command with all workbook operation subcommands.
    /// </summary>
    /// <returns>A Command object with Excel operation subcommands.</returns>
    public static Command GetCommand()
    {
        var command = new Command("excel", "Excel workbook operations.");

        command.Subcommands.Add(GetInfoCommand());
        command.Subcommands.Add(GetListSheetsCommand());
        command.Subcommands.Add(GetExtractSheetCommand());
        command.Subcommands.Add(GetToCsvCommand());
        command.Subcommands.Add(GetRefreshCommand());

        return command;
    }

    private static Command GetInfoCommand()
    {
        var command = new Command("info", "Display workbook metadata and information.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the Excel workbook file.",
        };

        command.Options.Add(inputOption);

        command.SetAction(parseResult => ShowInfo(parseResult.GetValue(inputOption)!));

        return command;
    }

    private static Command GetListSheetsCommand()
    {
        var command = new Command("list-sheets", "List all worksheet names in a workbook.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the Excel workbook file.",
        };

        command.Options.Add(inputOption);

        command.SetAction(parseResult => ListSheets(parseResult.GetValue(inputOption)!));

        return command;
    }

    private static Command GetExtractSheetCommand()
    {
        var command = new Command("extract-sheet", "Extract a specific sheet to a new workbook.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the Excel workbook file.",
        };

        var sheetOption = new Option<string>(name: "--sheet")
        {
            Description = "Name of the sheet to extract.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output workbook file.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(sheetOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            ExtractSheet(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(sheetOption)!,
                parseResult.GetValue(outputOption)!
            )
        );

        return command;
    }

    private static Command GetToCsvCommand()
    {
        var command = new Command("to-csv", "Convert workbook or sheet to CSV format.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the Excel workbook file.",
        };

        var sheetOption = new Option<string?>(name: "--sheet")
        {
            Description = "Name of the sheet to convert (optional, defaults to first sheet).",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output CSV file.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(sheetOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            ConvertToCsv(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                parseResult.GetValue(sheetOption)
            )
        );

        return command;
    }

    private static Command GetRefreshCommand()
    {
        var command = new Command("refresh", "Refresh all data connections in a workbook.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the Excel workbook file.",
        };

        command.Options.Add(inputOption);

        command.SetAction(parseResult => RefreshConnections(parseResult.GetValue(inputOption)!));

        return command;
    }

    internal static int ShowInfo(string inputPath)
    {
        try
        {
            Console.WriteLine($"Reading workbook: {inputPath}");

            var service = ServiceManager.GetService<ExcelService>();
            var info = service.GetWorkbookInfo(inputPath);

            Console.WriteLine("\n📊 Workbook Information:");
            Console.WriteLine($"  File: {info.FileName}");
            Console.WriteLine($"  Path: {info.Path}");
            Console.WriteLine($"  Sheets: {info.SheetCount}");
            Console.WriteLine($"  Author: {info.Author}");
            Console.WriteLine($"  Title: {info.Title}");
            Console.WriteLine($"  Subject: {info.Subject}");
            Console.WriteLine($"  Last Saved By: {info.LastSavedBy}");
            Console.WriteLine($"  Created: {info.Created:yyyy-MM-dd HH:mm:ss}");
            Console.WriteLine($"  Modified: {info.Modified:yyyy-MM-dd HH:mm:ss}");

            Console.WriteLine("\n📑 Worksheets:");
            for (int i = 0; i < info.SheetNames.Count; i++)
            {
                Console.WriteLine($"  {i + 1}. {info.SheetNames[i]}");
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }

    internal static int ListSheets(string inputPath)
    {
        try
        {
            Console.WriteLine($"Reading sheets from: {inputPath}");

            var service = ServiceManager.GetService<ExcelService>();
            var sheets = service.ListSheets(inputPath);

            Console.WriteLine($"\n📑 Found {sheets.Count} worksheet(s):");
            for (int i = 0; i < sheets.Count; i++)
            {
                Console.WriteLine($"  {i + 1}. {sheets[i]}");
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }

    internal static int ExtractSheet(string inputPath, string sheetName, string outputPath)
    {
        try
        {
            Console.WriteLine($"Extracting sheet '{sheetName}' from: {inputPath}");
            Console.WriteLine($"Output: {outputPath}");

            var service = ServiceManager.GetService<ExcelService>();
            service.ExtractSheet(inputPath, sheetName, outputPath);

            Console.WriteLine("✓ Sheet extracted successfully!");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }

    internal static int ConvertToCsv(string inputPath, string outputPath, string? sheetName)
    {
        try
        {
            Console.WriteLine($"Converting to CSV: {inputPath}");
            if (!string.IsNullOrEmpty(sheetName))
            {
                Console.WriteLine($"Sheet: {sheetName}");
            }
            Console.WriteLine($"Output: {outputPath}");

            var service = ServiceManager.GetService<ExcelService>();
            service.ConvertToCsv(inputPath, outputPath, sheetName);

            Console.WriteLine("✓ Conversion successful!");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }

    internal static int RefreshConnections(string inputPath)
    {
        try
        {
            Console.WriteLine($"Refreshing data connections: {inputPath}");

            var service = ServiceManager.GetService<ExcelService>();
            service.RefreshConnections(inputPath);

            Console.WriteLine("✓ Data connections refreshed successfully!");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }
}
