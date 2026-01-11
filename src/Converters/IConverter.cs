namespace msguru.Converters;

/// <summary>
/// Interface for document converters.
/// </summary>
public interface IConverter
{
    /// <summary>
    /// Converts a document from one format to another.
    /// </summary>
    /// <param name="inputPath">Path to the input file.</param>
    /// <param name="outputPath">Path to the output file.</param>
    /// <returns>The result of the conversion operation.</returns>
    ConversionResult Convert(string inputPath, string outputPath);
}
