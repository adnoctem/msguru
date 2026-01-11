namespace msguru.Converters;

/// <summary>
/// Represents the result of a document conversion operation.
/// </summary>
public class ConversionResult
{
    /// <summary>
    /// Indicates whether the conversion was successful.
    /// </summary>
    public bool Success { get; init; }

    /// <summary>
    /// The path to the output file if conversion was successful.
    /// </summary>
    public string? OutputPath { get; init; }

    /// <summary>
    /// Error message if conversion failed.
    /// </summary>
    public string? ErrorMessage { get; init; }

    /// <summary>
    /// Additional warnings or informational messages.
    /// </summary>
    public List<string> Messages { get; init; } = new();

    /// <summary>
    /// The number of items processed (e.g., pages, sheets, documents).
    /// </summary>
    public int ItemsProcessed { get; init; }

    /// <summary>
    /// Creates a successful conversion result.
    /// </summary>
    /// <param name="outputPath">The path to the output file.</param>
    /// <param name="itemsProcessed">The number of items processed (defaults to 1).</param>
    /// <returns>A ConversionResult indicating success.</returns>
    public static ConversionResult SuccessResult(string outputPath, int itemsProcessed = 1)
    {
        return new ConversionResult
        {
            Success = true,
            OutputPath = outputPath,
            ItemsProcessed = itemsProcessed
        };
    }

    /// <summary>
    /// Creates a failed conversion result.
    /// </summary>
    /// <param name="errorMessage">The error message describing the failure.</param>
    /// <returns>A ConversionResult indicating failure.</returns>
    public static ConversionResult FailureResult(string errorMessage)
    {
        return new ConversionResult
        {
            Success = false,
            ErrorMessage = errorMessage
        };
    }
}
