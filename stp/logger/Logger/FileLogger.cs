using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace logger.Logger
{
    public static class FileLogger
    {
        private static readonly NLog.Logger _logger = LogManager.GetCurrentClassLogger();

           public static void Debug(string message,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0)
        {
            _logger.Debug(FormatMessage(message, memberName, sourceFilePath, sourceLineNumber));
        }

        public static void Info(string message,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0)
        {
            _logger.Info(FormatMessage(message, memberName, sourceFilePath, sourceLineNumber));
        }

        public static void Warn(string message,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0)
        {
            _logger.Warn(FormatMessage(message, memberName, sourceFilePath, sourceLineNumber));
        }

        public static void Error(string message, Exception? ex = null,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0)
        {
            var msg = FormatMessage(message, memberName, sourceFilePath, sourceLineNumber);
            if (ex != null)
                _logger.Error(ex, msg);
            else
                _logger.Error(msg);
        }

        public static void Fatal(string message, Exception? ex = null,
            [CallerMemberName] string memberName = "",
            [CallerFilePath] string sourceFilePath = "",
            [CallerLineNumber] int sourceLineNumber = 0)
        {
            var msg = FormatMessage(message, memberName, sourceFilePath, sourceLineNumber);
            if (ex != null)
                _logger.Fatal(ex, msg);
            else
                _logger.Fatal(msg);
        }

       
        private static string FormatMessage(string message, string memberName, string filePath, int line)
        {
            if (string.IsNullOrEmpty(filePath))
                return message;

            var fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            return $"{message}  [at {fileName}.{memberName}:L{line}]";
        }
    }
}
