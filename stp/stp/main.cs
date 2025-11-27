using System.Runtime.InteropServices;
using System.IO;
using logger.Logger;
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
            Logger.Fatal("No doc found");
            return;
        }

        if (IsFileLocked.IsLocked(options.InputFile))
        {
            Logger.Fatal("Input file is locked. Close Word first.");
        }
        if (Path.Exists(options.OutputPath) && IsFileLocked.IsLocked(options.OutputPath))
        {
            Logger.Fatal("Output file is locked. Close Word first.");
        }

        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;

        Logger.Info("DoxC Conveer starting...");
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
        Logger.Info("DoxC Conveer complete.");

        if (Path.Exists(options.OutputPath))
        {
            Logger.Info("File found, OpenXML Conveer starting...");
            using (var doc = WordprocessingDocument.Open(options.OutputPath, isEditable: true))
            {
                OpenXMLConveer xmlConveer = new OpenXMLConveer(doc);
                xmlConveer.AllConveer();
                xmlConveer.Save();
            }
            Logger.Info("OpenXML Conveer complete.");
        }
        else
        {
            Logger.Info("Output file not found, OpenXML Conveer");
        }

        if (options.Rename != null)
        {
            options.OutputPath = FileRename.Rename(options.OutputPath, options.Rename!);
            bool renamed = true;
            if (!renamed)
            {
                Logger.Error("Cannot rename file");
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
                Logger.Warn($"Failed to open file - {ex.ToString()}");
            }
        }
    }
}