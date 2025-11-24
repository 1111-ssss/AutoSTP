namespace core.AppOptions
{
    public enum FileType
    {
        None = 0,
        LabWork = 1,
        PracticalWork = 2,
        GraduateWork = 3,
    }
    public class AppOptions
    {
        public String InputFile = "input.docx";
        public String OutputPath = "output.docx";
        public bool Verbose = false;
        public String? Author;
        public FileType FileType = FileType.None;
    }
}