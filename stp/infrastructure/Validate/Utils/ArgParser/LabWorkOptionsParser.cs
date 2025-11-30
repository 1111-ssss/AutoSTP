using core.Interfaces;
using core.Model;
using core.Enums;
using System.Text.Json;

namespace infrastructure.Validate.Utils.ArgParser
{
    public class LabWorkOptionsBuilder : IOptionsBuilder
    {
        public FileType TargetType => FileType.LabWork;

        public AppOptions BuildFromArgs(string[] args, AppOptions baseOptions)
        {
            var options = new LabWorkOptions
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
                if (args[i] == "--purpose" && i + 1 < args.Length)
                {
                    options.Purpose = args[++i];
                }
                else if (args[i] == "--labnumber" && i + 1 < args.Length)
                {
                    if (int.TryParse(args[++i], out int number))
                        options.LabNumber = number;
                    else
                        options.LabNumber = 1;
                }
                else if (args[i] == "--topic" && i + 1 < args.Length)
                {
                    options.Topic = args[++i];
                }
            }

            return options;
        }

        public AppOptions BuildFromJson(string json)
        {
            var opts = new JsonSerializerOptions { IncludeFields = true, PropertyNameCaseInsensitive = true, Converters = { new FileTypeConverter() } };
            return JsonSerializer.Deserialize<LabWorkOptions>(json, opts)
                   ?? throw new JsonException("Failed to parse LabWorkOptions from JSON");
        }
    }
}