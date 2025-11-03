using System;
using Compass.Infrastructure.Logging;
using Microsoft.Extensions.Configuration;
using NLog;

namespace Compass.Infrastructure.Logging;

public sealed class NLogLogger : ILog
{
    private readonly Logger _logger = LogManager.GetCurrentClassLogger();

    public NLogLogger()
    {
    }

    public NLogLogger(IConfiguration configuration)
    {
        if (configuration == null)
        {
            throw new ArgumentNullException(nameof(configuration));
        }
    }

    public void Info(string message) => _logger.Info(message);

    public void Warn(string message) => _logger.Warn(message);

    public void Error(string message, Exception? exception = null)
    {
        if (exception == null)
        {
            _logger.Error(message);
        }
        else
        {
            _logger.Error(exception, message);
        }
    }

    public void Debug(string message) => _logger.Debug(message);

    public IDisposable BeginScope(string scopeMessage)
    {
        return NestedDiagnosticsLogicalContext.Push(scopeMessage);
    }
}
