using Serilog;
using Serilog.Configuration;
using Serilog.Core;
using Serilog.Events;
using System;

namespace ExcelCharts2Powerpoint
{
    public class CustomSink : ILogEventSink
    {
        static public CustomSink instance;
        private readonly IFormatProvider _formatProvider;

        public delegate void LogEventHandler(string logText);
        public event LogEventHandler OnLogUpdate = delegate { };

        public CustomSink(IFormatProvider formatProvider)
        {
            instance = this;
            _formatProvider = formatProvider;
        }

        public void Emit(LogEvent logEvent)
        {
            var message = logEvent.RenderMessage(_formatProvider);
            if (OnLogUpdate != null) OnLogUpdate(DateTime.Now.ToString("[HH:mm:ss] ")+message);
        }
    }
    public static class CustomSinkExtensions
    {
        public static LoggerConfiguration CustomSinkLogger(
                  this LoggerSinkConfiguration loggerConfiguration,
                  IFormatProvider fmtProvider = null)
        {
            return loggerConfiguration.Sink(new CustomSink(fmtProvider));
        }
    }
}