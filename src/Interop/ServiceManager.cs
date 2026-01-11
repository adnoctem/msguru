using System.Runtime.InteropServices;

namespace msguru.Interop;

/// <summary>
/// Manages singleton instances of Office Interop services.
/// Ensures proper lifecycle management and COM object cleanup.
/// </summary>
public static class ServiceManager
{
    private static readonly object _lock = new();
    private static readonly Dictionary<Type, IOfficeService> _services = new();

    /// <summary>
    /// Gets or creates a singleton instance of the specified service type.
    /// </summary>
    /// <typeparam name="T">The service type to retrieve.</typeparam>
    /// <returns>The singleton instance of the service.</returns>
    public static T GetService<T>() where T : IOfficeService, new()
    {
        lock (_lock)
        {
            var type = typeof(T);
            if (!_services.ContainsKey(type))
            {
                var service = new T();
                _services[type] = service;
            }

            return (T)_services[type];
        }
    }

    /// <summary>
    /// Checks if a service instance exists for the specified type.
    /// </summary>
    public static bool HasService<T>() where T : IOfficeService
    {
        lock (_lock)
        {
            return _services.ContainsKey(typeof(T));
        }
    }

    /// <summary>
    /// Disposes and removes a specific service from the manager.
    /// </summary>
    public static void DisposeService<T>() where T : IOfficeService
    {
        lock (_lock)
        {
            var type = typeof(T);
            if (_services.TryGetValue(type, out var service))
            {
                service.Dispose();
                _services.Remove(type);
            }
        }
    }

    /// <summary>
    /// Disposes all managed services and clears the registry.
    /// Should be called on application shutdown.
    /// </summary>
    public static void DisposeAll()
    {
        lock (_lock)
        {
            foreach (var service in _services.Values)
            {
                try
                {
                    service.Dispose();
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error disposing service: {ex.Message}");
                }
            }

            _services.Clear();
        }
    }

    /// <summary>
    /// Releases a COM object and suppresses finalization.
    /// Helper method for proper COM cleanup.
    /// </summary>
    /// <param name="obj">The COM object to release.</param>
    public static void ReleaseComObject(object? obj)
    {
        if (obj != null && Marshal.IsComObject(obj))
        {
            try
            {
                if (OperatingSystem.IsWindows())
                {
                    Marshal.FinalReleaseComObject(obj);
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error releasing COM object: {ex.Message}");
            }
        }
    }
}
