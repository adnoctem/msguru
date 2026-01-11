using System.CommandLine;
using msguru.Converters;

namespace msguru.Commands;

/// <summary>
/// Provides commands for document format conversion operations.
/// </summary>
public static class ConvertCommand
{
    /// <summary>
    /// Gets the main convert command with all conversion subcommands.
    /// </summary>
    /// <returns>A Command object with conversion subcommands.</returns>
    public static Command GetCommand()
    {
        var command = new Command("convert", "Document conversion commands.");

        command.Subcommands.Add(GetDocxToHtmlCommand());
        command.Subcommands.Add(GetHtmlToDocxCommand());
        command.Subcommands.Add(GetXlsxToHtmlCommand());
        command.Subcommands.Add(GetHtmlToXlsxCommand());
        command.Subcommands.Add(GetDocxToPdfCommand());
        command.Subcommands.Add(GetXlsxToPdfCommand());
        command.Subcommands.Add(GetHtmlToPdfCommand());
        command.Subcommands.Add(GetTextToHtmlCommand());
        command.Subcommands.Add(GetHtmlToTextCommand());

        return command;
    }

    private static Command GetDocxToHtmlCommand()
    {
        var command = new Command("docx-to-html", "Convert DOCX document to HTML.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the input DOCX file.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output HTML file.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            ExecuteConversion(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                DocxConverter.ToHtml
            ));

        return command;
    }

    private static Command GetHtmlToDocxCommand()
    {
        var command = new Command("html-to-docx", "Convert HTML to DOCX document.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the input HTML file.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output DOCX file.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            ExecuteConversion(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                DocxConverter.FromHtml
            ));

        return command;
    }

    private static Command GetXlsxToHtmlCommand()
    {
        var command = new Command("xlsx-to-html", "Convert XLSX spreadsheet to HTML.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the input XLSX file.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output HTML file.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            ExecuteConversion(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                XlsxConverter.ToHtml
            ));

        return command;
    }

    private static Command GetHtmlToXlsxCommand()
    {
        var command = new Command("html-to-xlsx", "Convert HTML tables to XLSX spreadsheet.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the input HTML file.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output XLSX file.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            ExecuteConversion(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                XlsxConverter.FromHtml
            ));

        return command;
    }

    private static Command GetDocxToPdfCommand()
    {
        var command = new Command("docx-to-pdf", "Convert DOCX document to PDF.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the input DOCX file.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output PDF file.",
        };

        var chromePathOption = new Option<string?>(name: "--chrome-path")
        {
            Description = "Optional path to Chrome executable.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputOption);
        command.Options.Add(chromePathOption);

        command.SetAction(parseResult =>
            ExecuteAsyncConversion(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                parseResult.GetValue(chromePathOption),
                PdfConverter.DocxToPdfAsync
            ));

        return command;
    }

    private static Command GetXlsxToPdfCommand()
    {
        var command = new Command("xlsx-to-pdf", "Convert XLSX spreadsheet to PDF.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the input XLSX file.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output PDF file.",
        };

        var chromePathOption = new Option<string?>(name: "--chrome-path")
        {
            Description = "Optional path to Chrome executable.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputOption);
        command.Options.Add(chromePathOption);

        command.SetAction(parseResult =>
            ExecuteAsyncConversion(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                parseResult.GetValue(chromePathOption),
                PdfConverter.XlsxToPdfAsync
            ));

        return command;
    }

    private static Command GetHtmlToPdfCommand()
    {
        var command = new Command("html-to-pdf", "Convert HTML to PDF.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the input HTML file.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output PDF file.",
        };

        var chromePathOption = new Option<string?>(name: "--chrome-path")
        {
            Description = "Optional path to Chrome executable.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputOption);
        command.Options.Add(chromePathOption);

        command.SetAction(parseResult =>
            ExecuteAsyncConversion(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                parseResult.GetValue(chromePathOption),
                PdfConverter.HtmlToPdfAsync
            ));

        return command;
    }

    private static Command GetTextToHtmlCommand()
    {
        var command = new Command("text-to-html", "Convert plain text to HTML.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the input text file.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output HTML file.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            ExecuteConversion(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                HtmlConverter.FromPlainText
            ));

        return command;
    }

    private static Command GetHtmlToTextCommand()
    {
        var command = new Command("html-to-text", "Extract plain text from HTML.");

        var inputOption = new Option<string>(name: "--input")
        {
            Description = "Path to the input HTML file.",
        };

        var outputOption = new Option<string>(name: "--output")
        {
            Description = "Path to the output text file.",
        };

        command.Options.Add(inputOption);
        command.Options.Add(outputOption);

        command.SetAction(parseResult =>
            ExecuteConversion(
                parseResult.GetValue(inputOption)!,
                parseResult.GetValue(outputOption)!,
                HtmlConverter.ToPlainText
            ));

        return command;
    }

    internal static int ExecuteConversion(
        string inputPath,
        string outputPath,
        Func<string, string, ConversionResult> conversionFunc)
    {
        Console.WriteLine($"Converting: {inputPath}");
        Console.WriteLine($"Output: {outputPath}");

        var result = conversionFunc(inputPath, outputPath);

        if (result.Success)
        {
            Console.WriteLine($"✓ Conversion successful!");
            if (result.ItemsProcessed > 0)
            {
                Console.WriteLine($"  Processed {result.ItemsProcessed} item(s)");
            }
            foreach (var message in result.Messages)
            {
                Console.WriteLine($"  {message}");
            }
            return 0;
        }
        else
        {
            Console.Error.WriteLine($"✗ Conversion failed: {result.ErrorMessage}");
            return 1;
        }
    }

    internal static int ExecuteAsyncConversion(
        string inputPath,
        string outputPath,
        string? chromePath,
        Func<string, string, string?, Task<ConversionResult>> conversionFunc)
    {
        Console.WriteLine($"Converting: {inputPath}");
        Console.WriteLine($"Output: {outputPath}");

        var result = conversionFunc(inputPath, outputPath, chromePath).GetAwaiter().GetResult();

        if (result.Success)
        {
            Console.WriteLine($"✓ Conversion successful!");
            if (result.ItemsProcessed > 0)
            {
                Console.WriteLine($"  Processed {result.ItemsProcessed} item(s)");
            }
            foreach (var message in result.Messages)
            {
                Console.WriteLine($"  {message}");
            }
            return 0;
        }
        else
        {
            Console.Error.WriteLine($"✗ Conversion failed: {result.ErrorMessage}");
            return 1;
        }
    }
}
