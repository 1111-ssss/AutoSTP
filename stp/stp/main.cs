using System.Runtime.InteropServices;
using System.IO;
using logger.Logger;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using System.Diagnostics;
using core.Model;
using core.Enums;
using stp.Utils.FileUtils;
using stp.Utils.ArgParser;
using openXMlFunc.Conveers;
using DocXFunc.Conveers;
using infrastructure.OpenXmlContext;

class Program
{
    static void Main(String[] args)
    {
        AppOptions options = ArgParser.Parse(args);
        if (!Path.Exists(options.InputFile))
        {
            Logger.Log("No Doc found", LoggerState.Error);
            return;
        }
        if (IsFileLocked.IsLocked(options.InputFile))
        {
            Logger.Log("File is locked. Close Word first.", LoggerState.Error);
        }

        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;

        Logger.Log("DoxC Conveer starting...");
        DocXFunc.Conveers.LabsConveer doXcConveer = new DocXFunc.Conveers.LabsConveer(new infrastructure.Context.DocXContext(options.InputFile));
        doXcConveer.Conveer();
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

      
            Logger.Log("File found, OpenXML Conveer starting...");
            using (var doc = WordprocessingDocument.Open(options.OutputPath, isEditable: true))
            {
                var context = new DocContext(doc);
                openXMlFunc.Conveers.LabsConveer xmlConveer = new openXMlFunc.Conveers.LabsConveer(
                    context, 
                    new infrastructure.OpenXML.NamesProcessing.PicturesNames(context.Body)); //потом в picturesName добавить номер для нумеровки
                xmlConveer.Conveer();
                xmlConveer.Save();
            }
            Logger.Log("OpenXML Conveer complete.");
        
   



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
    }

}