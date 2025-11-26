using core.Model;
using core.Enums;
using logger.Logger;

namespace stp.Utils.ArgParser
{
    public class ArgParser
    {
        public static AppOptions Parse(String[] args)
        {
            var options = new AppOptions();

            var handlers = new Dictionary<string, Action<string>>
            {
                ["--input"] = val => options.InputFile = val,
                ["--output"] = val => options.OutputPath = val,
                ["--author"] = val => options.Author = val,
                ["--rename"] = val => options.Rename = val,
                ["--filetype"] = val =>
                {
                    if (Enum.TryParse<FileType>(val, true, out var type))
                        options.FileType = type;
                    else
                        throw new ArgumentException($"Неподдерживаемый тип документа: {val}. Допустимые: None, LabWork, PracticalWork, GraduateWork");
                },
            };

            for (int i = 0; i < args.Length; i++)
            {
                var arg = args[i];

                if (handlers.TryGetValue(arg, out var handler))
                {
                    if (i + 1 >= args.Length)
                        throw new ArgumentException($"Аргумент '{arg}' требует значение.");

                    handler(args[++i]);
                }
                else
                {
                    switch (arg)
                    {
                        case "--verbose":
                            options.Verbose = true;
                            break;
                        case "--open":
                            options.Open = true;
                            break;
                        case "--save":
                            options.Save = true;
                            break;
                        case "--log":
                            options.LoggerEnabled = true;
                            break;
                        default:
                            Logger.Detailed($"Неизвестный аргумент: {arg}", LoggerState.Warn);
                            break;
                            // throw new ArgumentException($"Неизвестный аргумент: {arg}");
                    }
                }
            }

            return options;
        }
    }
}