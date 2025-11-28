namespace core.Interfaces
{
    public interface IPipeline
    {
        void StartPipeline();
        void Save();
        void SaveAS(string TargetPath);
    }
}