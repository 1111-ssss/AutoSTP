using System.Runtime.InteropServices;
using System.IO;
using logger;
using DocumentFormat.OpenXml.Packaging;
using DocXFunc;
using openXMlFunc;
using System.Globalization;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using core;


class Program
{
    static void Main(String[] args)
    {
        Console.WriteLine("New7");
        
        String documentPath = args[0] ?? "input.docx";
        if (!Path.Exists(documentPath))
        {
            logger.Logger.Log("No doc found");
            return;
        }

        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;

        logger.Logger.Log("DoxC Conveer starting...");
        DoXcConveer doXcConveer = new DoXcConveer(documentPath);
        doXcConveer.AllConveer();
        doXcConveer.SaveAS("output.docx");
        logger.Logger.Log("DoxC Conveer complete.");

        if (Path.Exists("output.docx"))
        {
            logger.Logger.Log("Output file found, OpenXML Conveer starting...");
            using (var doc = WordprocessingDocument.Open("output.docx", isEditable: true))
            {
                OpenXMLConveer xmlConveer = new OpenXMLConveer(doc, stylesString: "Auto STP Script test by НЕГурский and НЕКасевич");
                xmlConveer.AllConveer();
                xmlConveer.Save();
            }
            logger.Logger.Log("OpenXML Conveer complete.");
        }
        else
        {
            logger.Logger.Log("Output file not found, OpenXML Conveer", loggerState: LoggerState.Error);
        }
        Console.ReadKey();
    }
    
}