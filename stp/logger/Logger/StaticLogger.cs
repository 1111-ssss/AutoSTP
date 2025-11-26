using System.Runtime.CompilerServices;
using core.Enums;
using core.Model;
using core.Interface;

namespace logger.Logger {
    public static class Logger
    {
        private static readonly ILogger _instance = new ConsoleLogger();
        private static LoggerState _minimalLogLevel = LoggerState.None;
        public static void ChangeLogLevel(LoggerState level)
        {
            _minimalLogLevel = level;
        }
        public static void Log(string message, LoggerState level = LoggerState.Info)
            => _instance.Log(message, level);

        public static void Detailed(
            string message,
            LoggerState level = LoggerState.Info,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string filePath = "",
            [CallerLineNumber] int lineNumber = 0
        )
        {
            _instance.Detailed(message, level, memberName, filePath, lineNumber);
        }
        public static void MapResult(TResult result) => _instance.MapResult(result);
        public static void Debug(String message) => Log(message, LoggerState.Debug);
        public static void Info(String message) => Log(message, LoggerState.Info);
        public static void Warn(String message) => Log(message, LoggerState.Warn);
        public static void Error(String message) => Log(message, LoggerState.Error);
        public static void Fatal(String message) => Log(message, LoggerState.Fatal);
    }
}