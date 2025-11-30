using core.Interfaces;
using core.Model;
using core.Enums;
using System.Text.Json;

namespace infrastructure.Validate.Utils.ArgParser
{
    public class PracticalWorkOptionsBuilder : IOptionsBuilder
    {
        public FileType TargetType => FileType.PracticalWork;

        public AppOptions BuildFromArgs(string[] args, AppOptions baseOptions)
        {
            var options = new PracticalWorkOptions
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
                else if (args[i] == "--pracnumber" && i + 1 < args.Length)
                {
                    if (int.TryParse(args[++i], out int number))
                        options.PracNumber = number;
                    else
                        options.PracNumber = 1;
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
            return JsonSerializer.Deserialize<PracticalWorkOptions>(json, opts)
                   ?? throw new JsonException("Failed to parse PracticalWorkOptions from JSON");
        }
    }
}