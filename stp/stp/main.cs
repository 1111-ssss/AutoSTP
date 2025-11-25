using System.Runtime.InteropServices;
using System.IO;
using logger;
using DocumentFormat.OpenXml.Packaging;
using DocXFunc;
using openXMlFunc;
using System.Globalization;
using core;


class Program
{
    static void Main(String[] args)
    {
        core.AppOptions.AppOptions options = core.ArgParser.ArgParser.Parse(args);
        if (!Path.Exists(options.InputFile))
        {
            Logger.Log("No doc found");
            return;
        }

        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;

        Logger.Log("DoxC Conveer starting...");
        DoXcConveer doXcConveer = new DoXcConveer(options.InputFile);
        doXcConveer.AllConveer();
        doXcConveer.SaveAS(options.OutputPath);
        Logger.Log("DoxC Conveer complete.");

        if (Path.Exists(options.OutputPath))
        {
            Logger.Log("Output file found, OpenXML Conveer starting...");
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
        Console.ReadKey();
    }
    
}