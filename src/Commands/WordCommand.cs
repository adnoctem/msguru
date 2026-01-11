using System.CommandLine;
using msguru.Interop;

namespace msguru.Commands;

/// <summary>
/// Provides commands for Word document operations using COM Interop.
/// </summary>
public static class WordCommand
{
    /// <summary>
    /// Gets the main Word command with all document operation subcommands.
    /// </summary>
    /// <returns>A Command object with Word operation subcommands.</returns>
    public static Command GetCommand()
    {
        var command = new Command("word", "Word document operations.");

        command.Subcommands.Add(GetInfoCommand());
        command.Subcommands.Add(GetExtractTextCommand());
        command.Subcommands.Add(GetExtractImagesCommand());
        command.Subcommands.Add(GetSearchReplaceCommand());
        command.Subcommands.Add(GetMergeCommand());

        return command;
    }

    private static Command GetInfoCommand()
    {
        var command = new Command("info", "Display document metadata and statistics.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the Word document file.",
        };

        command.Options.Add(inputOption);

        command.SetAction(parseResult => ShowInfo(parseResult.GetValue(inputOption)!));

        return command;
    }

    private static Command GetExtractTextCommand()
    {
        var command = new Command("extract-text", "Extract plain text content from a document.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the Word document file.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output text file.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            ExtractText(parseResult.GetValue(inputOption)!, parseResult.GetValue(outputOption)!)
        );

        return command;
    }

    private static Command GetExtractImagesCommand()
    {
        var command = new Command("extract-images", "Extract embedded images from a document.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the Word document file.",
        };

        var outputDirOption = new Option<string>(name: "--output-dir")
        {
            Description = "Directory to save extracted images.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputDirOption);

        command.SetAction(parseResult =>
            ExtractImages(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputDirOption)!
            )
        );

        return command;
    }

    private static Command GetSearchReplaceCommand()
    {
        var command = new Command("search-replace", "Find and replace text in a document.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the Word document file.",
        };

        var findOption = new Option<string>(name: "--find") { Description = "Text to find." };

        var replaceOption = new Option<string>(name: "--replace")
        {
            Description = "Text to replace with.",
        };

        var outputOption = new Option<string?>(name: "--output")
        {
            Description =
                "Path to save modified document (optional, modifies original if not specified).",
        };

        command.Options.Add(inputOption);
        command.Options.Add(findOption);
        command.Options.Add(replaceOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            SearchAndReplace(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(findOption)!,
                parseResult.GetValue(replaceOption)!,
                parseResult.GetValue(outputOption)
            )
        );

        return command;
    }

    private static Command GetMergeCommand()
    {
        var command = new Command("merge", "Merge multiple documents into one.");

        var inputsOption = new Option<string[]>(name: "--inputs")
        {
            Description = "Paths to the documents to merge (space-separated).",
            AllowMultipleArgumentsPerToken = true,
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output merged document.",
        };

        command.Options.Add(inputsOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            MergeDocuments(
                parseResult.GetValue(inputsOption)!.ToList(),
                parseResult.GetValue(outputOption)!
            )
        );

        return command;
    }

    internal static int ShowInfo(string inputPath)
    {
        try
        {
            Console.WriteLine($"Reading document: {inputPath}");

            var service = ServiceManager.GetService<WordService>();
            var info = service.GetDocumentInfo(inputPath);

            Console.WriteLine("\n📄 Document Information:");
            Console.WriteLine($"  File: {info.FileName}");
            Console.WriteLine($"  Path: {info.Path}");
            Console.WriteLine($"  Title: {info.Title}");
            Console.WriteLine($"  Author: {info.Author}");
            Console.WriteLine($"  Subject: {info.Subject}");
            Console.WriteLine($"  Keywords: {info.Keywords}");
            Console.WriteLine($"  Comments: {info.Comments}");
            Console.WriteLine($"  Last Saved By: {info.LastSavedBy}");
            Console.WriteLine($"  Created: {info.Created:yyyy-MM-dd HH:mm:ss}");
            Console.WriteLine($"  Modified: {info.Modified:yyyy-MM-dd HH:mm:ss}");

            Console.WriteLine("\n📊 Statistics:");
            Console.WriteLine($"  Pages: {info.PageCount}");
            Console.WriteLine($"  Words: {info.WordCount}");
            Console.WriteLine($"  Characters: {info.CharacterCount}");
            Console.WriteLine($"  Paragraphs: {info.ParagraphCount}");

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }

    internal static int ExtractText(string inputPath, string outputPath)
    {
        try
        {
            Console.WriteLine($"Extracting text from: {inputPath}");
            Console.WriteLine($"Output: {outputPath}");

            var service = ServiceManager.GetService<WordService>();
            var text = service.ExtractText(inputPath);

            File.WriteAllText(outputPath, text);

            Console.WriteLine($"✓ Extracted {text.Length} characters!");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }

    internal static int ExtractImages(string inputPath, string outputDir)
    {
        try
        {
            Console.WriteLine($"Extracting images from: {inputPath}");
            Console.WriteLine($"Output directory: {outputDir}");

            var service = ServiceManager.GetService<WordService>();
            var extractedFiles = service.ExtractImages(inputPath, outputDir);

            Console.WriteLine($"✓ Extracted {extractedFiles.Count} image(s)!");
            foreach (var file in extractedFiles)
            {
                Console.WriteLine($"  - {file}");
            }

            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }

    internal static int SearchAndReplace(
        string inputPath,
        string findText,
        string replaceText,
        string? outputPath
    )
    {
        try
        {
            Console.WriteLine($"Searching and replacing in: {inputPath}");
            Console.WriteLine($"Find: '{findText}'");
            Console.WriteLine($"Replace with: '{replaceText}'");
            if (!string.IsNullOrEmpty(outputPath))
            {
                Console.WriteLine($"Output: {outputPath}");
            }

            var service = ServiceManager.GetService<WordService>();
            var count = service.SearchAndReplace(inputPath, findText, replaceText, outputPath);

            Console.WriteLine($"✓ Replaced {count} occurrence(s)!");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }

    internal static int MergeDocuments(List<string> inputPaths, string outputPath)
    {
        try
        {
            Console.WriteLine($"Merging {inputPaths.Count} document(s):");
            foreach (var path in inputPaths)
            {
                Console.WriteLine($"  - {path}");
            }
            Console.WriteLine($"Output: {outputPath}");

            var service = ServiceManager.GetService<WordService>();
            service.MergeDocuments(inputPaths, outputPath);

            Console.WriteLine("✓ Documents merged successfully!");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"✗ Error: {ex.Message}");
            return 1;
        }
    }
}
