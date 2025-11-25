using System.Runtime.InteropServices;
using System.IO;
using logger;
using DocumentFormat.OpenXml.Packaging;
using DocXFunc;
using openXMlFunc;
using System.Globalization;
using System.Diagnostics;
using core.Model;
using core.Enums;
using stp.Utils.FileUtils;
using stp.Utils.ArgParser;

class Program
{
    static void Main(String[] args)
    {
        AppOptions options = ArgParser.Parse(args);
        if (!Path.Exists(options.InputFile))
        {
            Logger.Log("No doc found", LoggerState.Error);
            return;
        }
        if (IsFileLocked.IsLocked(options.InputFile))
        {
            Logger.Log("File is locked. Close Word first.", LoggerState.Error);
        }

        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;

        Logger.Log("DoxC Conveer starting...");
        DoXcConveer doXcConveer = new DoXcConveer(options.InputFile);
        doXcConveer.AllConveer();
        if (options.Save)
        {
            options.OutputPath = options.InputFile;
            doXcConveer.Save();
        }
        else
        {
            doXcConveer.SaveAS(options.OutputPath);
        }
        Logger.Log("DoxC Conveer complete.");

        if (Path.Exists(options.OutputPath))
        {
            Logger.Log("File found, OpenXML Conveer starting...");
            using (var doc = WordprocessingDocument.Open(options.OutputPath, isEditable: true))
            {
                OpenXMLConveer xmlConveer = new OpenXMLConveer(doc);
                xmlConveer.AllConveer();
                xmlConveer.Save();
            }
            Logger.Log("OpenXML Conveer complete.");
        }
        else
        {
            Logger.Log("Output file not found, OpenXML Conveer", loggerState: LoggerState.Error);
        }
        Console.WriteLine(options.Rename);
        if (options.Rename != null)
        {
            options.OutputPath = FileRename.Rename(options.OutputPath, options.Rename!);
            bool renamed = true;
            if (!renamed)
            {
                Logger.Log("Cannot rename file", LoggerState.Error);
            }
        }
        if (options.Open)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = options.OutputPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                Logger.Log($"Failed to open file - {ex.ToString()}", LoggerState.Error);
            }
        }
        Console.ReadKey();
    }

}