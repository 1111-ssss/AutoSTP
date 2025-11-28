using core.Model;
using logger.Logger;
using infrastructure.Validate.Utils.FileUtils;
using System.Diagnostics;

namespace infrastructure.Validate
{
    public class ValidateArgs
    {
        public static bool ValidateOptions(AppOptions options)
        {
            if (!Path.Exists(options.InputFile))
            {
                Logger.Fatal("No doc found");
                return false;
            }
            if (IsFileLocked.IsLocked(options.InputFile))
            {
                Logger.Fatal("Input file is locked. Close Word first.");
                return false;
            }
            if (Path.Exists(options.OutputPath) && IsFileLocked.IsLocked(options.OutputPath))
            {
                Logger.Fatal("Output file is locked. Close Word first.");
                return false;
            }

            return true;
        }
        public static void OpenFile(String pathToFile) {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = pathToFile,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Logger.Warn($"Failed to open file - {ex.ToString()}");
            }
        }
        public static void RenameFile(ref AppOptions options) {
            if (options.Rename != null)
            {
                options.OutputPath = FileRename.Rename(options.OutputPath, options.Rename!);
                bool renamed = true;
                if (!renamed)
                {
                    Logger.Error("Cannot rename file");
                }
            }
        }
    }
}