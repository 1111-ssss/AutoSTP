using core.Model;
using core.Enums;
using logger.Logger;
using System.Net;
using System.Text.Json;

namespace infrastructure.Validate.Utils.ArgParser
{
    public class ArgParser
    {
        private static readonly Dictionary<FileType, IOptionsBuilder> _builders = new()
        {
            { FileType.LabWork, new LabWorkOptionsBuilder() },
            { FileType.PracticalWork, new PracticalWorkOptionsBuilder() },
            { FileType.GraduateWork, new GraduateWorkOptionsBuilder() }
        };
        public static AppOptions Parse(String[] args)
        {
            if (args.Length == 1 && File.Exists(args[0]) && Path.GetExtension(args[0]).Equals(".json", StringComparison.OrdinalIgnoreCase))
            {
                var json = File.ReadAllText(args[0]);
                // Сначала прочитаем fileType из JSON
                using var doc = JsonDocument.Parse(json);
                if (!doc.RootElement.TryGetProperty("FileType", out var typeElement))
                    throw new ArgumentException("JSON должен содержать 'FileType'.");

                var typeName = typeElement.GetString();
                if (string.IsNullOrEmpty(typeName) || !Enum.TryParse<FileType>(typeName, true, out var fileType))
                    throw new ArgumentException($"Неверный FileType в JSON: {typeName}");

                if (_builders.TryGetValue(fileType, out var optionsBuilder))
                    return optionsBuilder.BuildFromJson(json);

                throw new NotSupportedException($"Поддержка {fileType} не реализована.");
            }

            // return ParseInternal(args);
            var baseOptions = ParseCommonArgs(args, out var remainingArgs);

            if (baseOptions.FileType == FileType.None)
                throw new ArgumentException("Требуется указать --filetype.");

            if (_builders.TryGetValue(baseOptions.FileType, out var builder))
                return builder.BuildFromArgs(remainingArgs, baseOptions);

            return baseOptions; // fallback (но лучше не использовать)
        }
        private static AppOptions ParseCommonArgs(String[] args, out string[] remainingArgs)
        {
            var baseOptions = new AppOptions();
            var usedIndices = new HashSet<int>();

            var handlers = new Dictionary<string, Action<string>>
            {
                ["--input"] = v => { baseOptions.InputFile = v; },
                ["--output"] = v => { baseOptions.OutputPath = v; },
                ["--author"] = v => { baseOptions.Author = v; },
                ["--rename"] = v => { baseOptions.Rename = v; },
                ["--filetype"] = v =>
                {
                    if (Enum.TryParse<FileType>(v, true, out var t))
                        baseOptions.FileType = t;
                }
            };


            for (int i = 0; i < args.Length; i++)
            {
                var arg = args[i];

                if (handlers.TryGetValue(arg, out var handler) && i + 1 >= args.Length)
                {
                    handler(args[++i]);
                    usedIndices.Add(i);
                    usedIndices.Add(i + 1);
                    i++;
                }
                else
                {
                    switch (arg)
                    {
                        case "--verbose":
                            baseOptions.Verbose = true;
                            usedIndices.Add(i);
                            break;
                        case "--open":
                            baseOptions.Open = true;
                            usedIndices.Add(i);
                            break;
                        case "--save":
                            baseOptions.Save = true;
                            usedIndices.Add(i);
                            break;
                        case "--nolog":
                            baseOptions.LoggerEnabled = false;
                            usedIndices.Add(i);
                            break;
                        default:
                            Logger.Warn($"Unknown argument {arg}");
                            break;
                    }
                }
            }
            remainingArgs = args.Where((_, i) => !usedIndices.Contains(i)).ToArray();
            return baseOptions;
        }
    }
}