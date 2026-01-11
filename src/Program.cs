using System.CommandLine;
using System.Runtime.InteropServices;
using msguru.Commands;
using msguru.Interop;

namespace msguru;

class Program
{
    static async Task<int> Main(string[] args)
    {
        var rootCommand = new RootCommand("Msguru CLI tool for messaging utilities.");

        // Add Office Interop commands (Windows only)
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            rootCommand.Subcommands.Add(OutlookCommand.GetCommand());
            rootCommand.Subcommands.Add(WordCommand.GetCommand());
            rootCommand.Subcommands.Add(ExcelCommand.GetCommand());
        }
        else
        {
            // Add placeholder commands with helpful error messages for non-Windows platforms
            rootCommand.Subcommands.Add(
                CreateUnsupportedCommand("outlook", "Outlook management commands.")
            );
            rootCommand.Subcommands.Add(
                CreateUnsupportedCommand("word", "Word document operations.")
            );
            rootCommand.Subcommands.Add(
                CreateUnsupportedCommand("excel", "Excel workbook operations.")
            );
        }

        // Add cross-platform convert commands
        rootCommand.Subcommands.Add(ConvertCommand.GetCommand());

        var result = await rootCommand.Parse(args).InvokeAsync();

        // Cleanup all Office services on exit (Windows only)
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            ServiceManager.DisposeAll();
        }

        return result;
    }

    private static Command CreateUnsupportedCommand(string name, string description)
    {
        var command = new Command(name, description);

        command.SetAction(_ =>
        {
            var platform = RuntimeInformation.IsOSPlatform(OSPlatform.OSX) ? "macOS" : "Linux";
            Console.Error.WriteLine($"✗ The '{name}' command is only available on Windows.");
            Console.Error.WriteLine(
                $"  Microsoft Office COM Interop is not supported on {platform}."
            );
            Console.Error.WriteLine();
            Console.Error.WriteLine($"  Available commands on {platform}:");
            Console.Error.WriteLine(
                $"    - convert: Document format conversion (DOCX, XLSX, HTML, PDF)"
            );
            return 1;
        });

        return command;
    }
}
