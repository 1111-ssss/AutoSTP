using core.Interfaces;
using core.Model;
using core.Enums;
using System.Text.Json;

namespace infrastructure.Validate.Utils.ArgParser
{
    public class GraduateWorkOptionsBuilder : IOptionsBuilder
    {
        public FileType TargetType => FileType.GraduateWork;

        public AppOptions BuildFromArgs(string[] args, AppOptions baseOptions)
        {
            var options = new GraduateWorkOptions
            {
                InputFile = baseOptions.InputFile,
                OutputPath = baseOptions.OutputPath,
                Author = baseOptions.Author,
                Rename = baseOptions.Rename,
                FileType = baseOptions.FileType,
                Verbose = baseOptions.Verbose,
                Open = baseOptions.Open,
                Save = baseOptions.Save,
                LoggerEnabled = baseOptions.LoggerEnabled
            };

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "--topic" && i + 1 < args.Length)
                {
                    options.Topic = args[++i];
                }
                else if (args[i] == "--groupnumber" && i + 1 < args.Length)
                {
                    if (int.TryParse(args[++i], out int number))
                        options.GroupNumber = number;
                    else
                        options.GroupNumber = 1;
                }
                else if (args[i] == "--prepod" && i + 1 < args.Length)
                {
                    options.Prepod = args[++i];
                }
            }

            return options;
        }

        public AppOptions BuildFromJson(string json)
        {
            var opts = new JsonSerializerOptions { IncludeFields = true, PropertyNameCaseInsensitive = true, Converters = { new FileTypeConverter() } };
            return JsonSerializer.Deserialize<GraduateWorkOptions>(json, opts)
                   ?? throw new JsonException("Failed to parse GraduateWorkOptions from JSON");
        }
    }
}