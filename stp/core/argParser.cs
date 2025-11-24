namespace core.ArgParser
{
    public class ArgParser
    {
        public static core.AppOptions.AppOptions Parse(String[] args)
        {
            var options = new core.AppOptions.AppOptions();

            var handlers = new Dictionary<string, Action<string>>
            {
                ["--input"] = val => options.InputFile = val,
                ["--output"] = val => options.OutputPath = val,
                ["--author"] = val => options.Author = val,
                ["--filetype"] = val =>
                {
                    if (Enum.TryParse<core.AppOptions.FileType>(val, true, out var type))
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
                        default:
                            throw new ArgumentException($"Неизвестный аргумент: {arg}");
                    }
                }
            }

            return options;
        }
    }
}