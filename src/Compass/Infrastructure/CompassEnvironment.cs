using System;
using System.Text;
using Compass.Infrastructure.Logging;
using Compass.Models;
using Microsoft.Extensions.Configuration;

namespace Compass.Infrastructure;

/// <summary>
/// Centralized bootstrapper for configuration, logging, and miscellaneous environment setup.
/// </summary>
public static class CompassEnvironment
{
    private static readonly object SyncRoot = new();
    private static bool _initialized;
    private static IConfigurationRoot? _configuration;
    private static AppSettings _appSettings = new();
    private static ILog _log = new NLogLogger();

    /// <summary>
    /// Ensures environment initialization occurs exactly once.
    /// </summary>
    public static void Initialize()
    {
        if (_initialized)
        {
            return;
        }

        lock (SyncRoot)
        {
            if (_initialized)
            {
                return;
            }

            try
            {
                var providerType = Type.GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages", throwOnError: false);
                if (providerType?.GetProperty("Instance")?.GetValue(null) is EncodingProvider provider)
                {
                    Encoding.RegisterProvider(provider);
                }
            }
            catch (Exception ex)
            {
                _log.Warn($"Failed to register code page provider: {ex.Message}");
            }

            try
            {
                _configuration = new ConfigurationBuilder()
                    .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: false)
                    .Build();

                var settings = new AppSettings();
                _configuration.Bind(settings);
                _appSettings = settings;

                _log = new NLogLogger(_configuration);
            }
            catch (Exception ex)
            {
                _log.Warn($"Failed to load configuration: {ex.Message}");
                _configuration = null;
                _appSettings = new AppSettings();
            }

            _initialized = true;
        }
    }

    public static IConfiguration Configuration
        => (_configuration ?? new ConfigurationBuilder().Build());

    public static AppSettings AppSettings => _appSettings;

    public static ILog Log => _log;
}
