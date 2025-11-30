using core.Enums;
using core.Model;

public interface IOptionsBuilder
{
    FileType TargetType { get; }
    AppOptions BuildFromArgs(string[] args, AppOptions baseOptions);
    AppOptions BuildFromJson(string json);
}