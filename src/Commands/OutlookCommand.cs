using System.CommandLine;
using msguru.Interop;

namespace msguru.Commands;

/// <summary>
/// Provides commands for Outlook email and PST management operations using COM Interop.
/// </summary>
public static class OutlookCommand
{
    /// <summary>
    /// Gets the main Outlook command with all email management subcommands.
    /// </summary>
    /// <returns>A Command object with Outlook operation subcommands.</returns>
    public static Command GetCommand()
    {
        var command = new Command("outlook", "Outlook management commands.");
        command.Subcommands.Add(GetArchiveCommand());
        command.Subcommands.Add(GetListFoldersCommand());
        command.Subcommands.Add(GetExportCommand());
        return command;
    }

    private static Command GetArchiveCommand()
    {
        var command = new Command("archive", "Archive outlook items based on criteria.");

        var daysOption = new Option<int>(name: "--days")
        {
            Description = "Number of days to keep before archiving.",
            DefaultValueFactory = _ => 30,
        };

        var pathOption = new Option<string?>(name: "--path")
        {
            Description =
                "Path to the folder to archive. If not specified, all folders are considered.",
        };

        var dryRunOption = new Option<bool>(name: "--dry-run")
        {
            Description = "Perform a dry run without making any changes.",
            DefaultValueFactory = _ => false,
        };

        command.Options.Add(daysOption);
        command.Options.Add(pathOption);
        command.Options.Add(dryRunOption);

        command.SetAction(parseResult =>
            (
                Archive(
                    parseResult.GetValue(daysOption),
                    parseResult.GetValue(pathOption),
                    parseResult.GetValue(dryRunOption)
                )
            )
        );

        return command;
    }

    private static Command GetListFoldersCommand()
    {
        var command = new Command("list-folders", "List all Outlook folders.");

        command.SetAction(parseResult => ListFolders());

        return command;
    }

    private static Command GetExportCommand()
    {
        var command = new Command("export", "Export messages to MSG files.");

        var folderOption = new Option<string>(name: "--folder")
        {
            Description = "Folder path to export from.",
            DefaultValueFactory = _ => "Inbox",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Output directory for exported messages.",
        };

        var maxOption = new Option<int?>(name: "--max")
        {
            Description = "Maximum number of messages to export (optional).",
        };

        command.Options.Add(folderOption);
        command.Options.Add(outputOption);
        command.Options.Add(maxOption);

        command.SetAction(parseResult =>
            ExportMessages(
                parseResult.GetValue(folderOption)!,
                parseResult.GetValue(outputOption)!,
                parseResult.GetValue(maxOption)
            )
        );

        return command;
    }

    internal static int Archive(int days, string? path, bool dryRun)
    {
        try
        {
            Console.WriteLine($"Running Outlook Archive...");
            Console.WriteLine($"Criteria: Older than {days} days");
            Console.WriteLine($"Target Path: {path ?? "All Folders"}");

            var service = ServiceManager.GetService<OutlookService>();
            var result = service.ArchiveItems(days, path, dryRun);

            Console.WriteLine($"\n📊 Results:");
            Console.WriteLine($"  Folders processed: {result.ProcessedFolders}");
            Console.WriteLine($"  Items found: {result.ItemsFound}");

            if (dryRun)
            {
                Console.WriteLine($"\n[Dry Run] Would archive {result.ItemsFound} item(s):");
                foreach (var item in result.ArchivedItems.Take(10))
                {
                    Console.WriteLine($"  - {item}");
                }
                if (result.ArchivedItems.Count > 10)
                {
                    Console.WriteLine($"  ... and {result.ArchivedItems.Count - 10} more");
                }
            }
            else
            {
                Console.WriteLine($"  Items archived: {result.ItemsArchived}");
            }

            if (result.Errors.Count > 0)
            {
                Console.Error.WriteLine($"\n⚠ Errors encountered:");
                foreach (var error in result.Errors)
                {
                    Console.Error.WriteLine($"  - {error}");
                }
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }

    internal static int ListFolders()
    {
        try
        {
            Console.WriteLine("Listing Outlook folders...\n");

            var service = ServiceManager.GetService<OutlookService>();
            var folders = service.ListFolders();

            Console.WriteLine($"📁 Found {folders.Count} folder(s):");
            foreach (var folder in folders)
            {
                Console.WriteLine($"  {folder}");
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }

    internal static int ExportMessages(string folder, string outputDir, int? maxCount)
    {
        try
        {
            Console.WriteLine($"Exporting messages from: {folder}");
            Console.WriteLine($"Output directory: {outputDir}");
            if (maxCount.HasValue)
            {
                Console.WriteLine($"Max messages: {maxCount.Value}");
            }

            var service = ServiceManager.GetService<OutlookService>();
            var exportCount = service.ExportMessages(folder, outputDir, maxCount);

            Console.WriteLine($"✓ Exported {exportCount} message(s)!");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }
}
