using System;

namespace VUS_Assist_v._1._0
{
    public class LogEntry
    {
        public DateTime Timestamp { get; set; }
        public string Tag { get; set; }
        public string Message { get; set; }
        public LogLevel Level { get; set; }

        public LogEntry(LogLevel level, string tag, string message)
        {
            Timestamp = DateTime.Now;
            Level = level;
            Tag = tag;
            Message = message;
        }

        public string FormattedMessage => $"[{Timestamp:HH:mm:ss.fff}] [{Tag}] {Message}";
    }

    public enum LogLevel
    {
        Info,
        Warning,
        Error,
        Success,
        Database,
        Generation,
        UI
    }
}