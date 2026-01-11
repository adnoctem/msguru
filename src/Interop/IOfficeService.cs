using System.Runtime.InteropServices;

namespace msguru.Interop;

/// <summary>
/// Base interface for all Office COM Interop services.
/// </summary>
public interface IOfficeService : IDisposable
{
    /// <summary>
    /// Gets the name of the Office application (e.g., "Excel", "Word", "Outlook").
    /// </summary>
    string ApplicationName { get; }

    /// <summary>
    /// Indicates whether the Office application is currently running.
    /// </summary>
    bool IsApplicationRunning { get; }

    /// <summary>
    /// Initializes the COM Interop connection to the Office application.
    /// </summary>
    void Initialize();

    /// <summary>
    /// Cleans up COM objects and releases resources.
    /// </summary>
    void Cleanup();
}
