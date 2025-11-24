using System.Collections;
using System.Threading.Tasks.Sources;
using core;

namespace logger
{
    public class Logger
    {
        public static void Log(String err, LoggerState loggerState = LoggerState.Info)
        {
            switch (loggerState)
            {
                case LoggerState.Info:
                    Console.WriteLine($"Information: {err}");
                    break;
                case LoggerState.Warn:
                    Console.WriteLine($"Warn: {err}");
                    break;
                case LoggerState.Error:
                    Console.WriteLine($"Error: {err}");
                    break;
            }
        }
    }
}