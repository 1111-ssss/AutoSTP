using core;
using core.Enums;
using infrastructure.Context;
using logger.Logger;
using infrastructure;
using infrastructure.OpenXML.Metatags;
using infrastructure.OpenXML.Style;
using infrastructure.OpenXML.Margins;
using infrastructure.OpenXML.NamesProcessing;
using core.Interfaces;
using DocumentFormat.OpenXml.InkML;
using core.Model;

namespace application.Pipelines.OpenXML
{
    public class LabWorkOpenXMLPipeline : IPipeline
    {
        private OpenXmlContext _context;
        private PicturesNames _picturesNames;

        public LabWorkOpenXMLPipeline(OpenXmlContext doc, LabWorkOptions options)
        {
            _context = doc;
            // _picturesNames = picturesNames;
            _picturesNames = new PicturesNames(_context.Body, 1);
        }

        public void StartPipeline()
        {
            try
            {
                Margins.SetupPageMargins(_context.Body);
                Logger.Log("OpenXML Conveer: SetupPageMargins is done");
                EditStyle.CreateBaseStyle(_context);
                Logger.Log("OpenXML Conveer: CreateBaseStyle is done");
                _picturesNames.AddPictureNames();
                TableNames.TableName(_context.Body);
                TableStyle.ApplyBaseTableStyle(_context.Body);
                TextStyle.ApplyBaseTextStyle(_context.Body);
                Logger.Log("OpenXML Conveer: ApplyBaseTextStyle is done");
                Logger.Log("OpenXML Conveer: TableNames is done");
                Metatags.AddMetatags(_context, author: "AutoSTP", description: "Document formatted with private docx formatter script");
                Logger.Log("OpenXML Conveer: AddMultipleStyles is done");
                StylesCreator.AddMultipleStyles(_context.MainDocumentPart, new Dictionary<string, string>() {
                    {"stp1", "1 Automatical"},
                    {"stp2", "2 STP"},
                    {"stp3", "3 Formatter"},
                    {"stp4", "4 Script"},
                });
                Logger.Log("OpenXML Conveer: AddMetatags is done");
            }
            catch (Exception ex)
            {
                Logger.Log($"Formatting - {ex.ToString()}", LoggerState.Error);
                return;
                // return false;
            }
            finally
            {
                Logger.Log("OpenXML: file formatted.");
            }
            return;
            // return true;
        }
        public void Save()
        {
            try
            {
                _context.Doc.Save();
            }
            catch (Exception ex)
            {
                Logger.Log($"Saving - {ex.ToString()}", LoggerState.Error);
                return;
            }
            finally
            {
                Logger.Log("OpenXML: file saved.");
            }
            return;
        }
        public void SaveAS(String TargetPath) => Save();
    }
}