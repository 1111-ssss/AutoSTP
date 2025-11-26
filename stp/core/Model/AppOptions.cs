using core.Enums;

namespace core.Model
{
    public class AppOptions
    {
        public String InputFile = "input.docx";
        public String OutputPath = "output.docx";
        public bool Verbose = false;
        public String? Author;
        public FileType FileType = FileType.None;
        public bool Open = false;
        public bool Save = false;
        public String? Rename;
        public bool LoggerEnabled = false;
    }
}