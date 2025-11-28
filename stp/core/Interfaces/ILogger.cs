using System.Runtime.CompilerServices;
using core.Model;
using core.Enums;

namespace core.Interfaces
{
    public interface ILogger
    {
        void Log(String message, LoggerState level = LoggerState.Info);
        void Debug(String message);
        void Info(String message);
        void Warn(String message);
        void Error(String message);
        void Fatal(String message);
        void Detailed(
            string message,
            LoggerState level = LoggerState.Info,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string filePath = "",
            [CallerLineNumber] int lineNumber = 0
        );
        void MapResult(TResult result); // удалить НАЗУЙ
    }
}