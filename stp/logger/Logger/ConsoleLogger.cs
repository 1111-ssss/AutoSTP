using System.Runtime.CompilerServices;
using core.Interfaces;
using core.Enums;
using core.Model;

namespace logger.Logger
{
    public class ConsoleLogger : ILogger
    {
        private LoggerState _minimalLogLevel = LoggerState.None;
        public void Log(String message, LoggerState level = LoggerState.Info)
        {
            if (level > _minimalLogLevel && level == LoggerState.None)
            {
                return;
            }
            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            var levelStr = level switch
            {
                LoggerState.Debug => "DEBUG",
                LoggerState.Info => "INFO",
                LoggerState.Warn => "WARN",
                LoggerState.Error => "ERROR",
                LoggerState.Fatal => "FATAL",
                _ => "DEBUG",
            };
            var formatted = $"[{timestamp}] [{levelStr}] {message}";

            ConsoleColor originalColor = Console.ForegroundColor;
            try
            {
                Console.ForegroundColor = level switch
                {
                    LoggerState.Debug => originalColor,
                    LoggerState.Info => ConsoleColor.Gray,
                    LoggerState.Warn => ConsoleColor.Yellow,
                    LoggerState.Error => ConsoleColor.Red,
                    LoggerState.Fatal => ConsoleColor.Red,
                    _ => originalColor
                };

                Console.WriteLine(formatted);
            }
            finally
            {
                Console.ForegroundColor = originalColor;
            }
        }
        public void Detailed(
            string message,
            LoggerState level = LoggerState.Info,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string filePath = "",
            [CallerLineNumber] int lineNumber = 0)
        {
            var fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            Log($"{fileName}.{memberName}():{lineNumber} → {message}", level);
        }
        public void MapResult(TResult result)
        {
            if (!result.IsCompleted)
            {
                int errNum = (int)result.ErrorCode;
                String errMessage = GetErrorMessage(result.ErrorCode);
                if (errNum >= 200 && errNum < 300)
                {
                    Warn(errMessage);
                }
                else if (errNum >= 300 && errNum < 400)
                {
                    Error(errMessage);
                }
                else
                {
                    Fatal(errMessage);
                }
            }
        }
        private static String GetErrorMessage(errorCode code)
        {
            return code switch
            {
                // errorCode.ValidationError => "Validation failed",
                // errorCode.InternalError => "Unexpected internal error",
                _ => $"Unknown error: {code}"
            };
        }
        public void Debug(String message) => Log(message, LoggerState.Debug);
        public void Info(String message) => Log(message, LoggerState.Info);
        public void Warn(String message) => Log(message, LoggerState.Warn);
        public void Error(String message) => Log(message, LoggerState.Error);
        public void Fatal(String message) => Log(message, LoggerState.Fatal);
    }
}