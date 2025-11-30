using core.Enums;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace infrastructure.Validate.Utils.ArgParser
{
    public class FileTypeConverter : JsonConverter<FileType>
    {
        public override FileType Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            var value = reader.GetString()?.Trim();
            if (string.IsNullOrEmpty(value))
                return FileType.None;

            if (Enum.TryParse<FileType>(value, ignoreCase: true, out var result))
                return result;

            throw new JsonException($"Неизвестный FileType: '{value}'. Допустимые: LabWork, PracticalWork, GraduateWork");
        }

        public override void Write(Utf8JsonWriter writer, FileType value, JsonSerializerOptions options)
        {
            writer.WriteStringValue(value.ToString());
        }
    }
}