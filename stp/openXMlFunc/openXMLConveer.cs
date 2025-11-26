using core;
using DocumentFormat.OpenXml.Packaging;
using logger.Logger;
using core.Enums;
using openXMlFunc.Style;
namespace openXMlFunc
{
    public class OpenXMLConveer
    {
        private WordprocessingDocument _doc;
        public OpenXMLConveer(WordprocessingDocument doc)
        {
            _doc = doc;
        }




        public bool AllConveer()
        {
            try
            {
                Style.Margins.Margins.SetupPageMargins(_doc);
                Logger.Log("OpenXML Conveer: SetupPageMargins is done");
                // Style.Style.EditStyle.CreateBaseStyle(_doc);
                // logger.Logger.Log("OpenXML Conveer: CreateBaseStyle is done");
                Style.Style.TextStyle.ApplyBaseStyle(_doc);
                Logger.Log("OpenXML Conveer: ApplyBaseStyle is done");
                Style.Names.TableNames.TableName(_doc, 1);
                Logger.Log("OpenXML Conveer: TableNames is done");
                Style.Metatags.Metatags.AddMetatags(_doc, author: "AutoSTP", description: "Document formatted with private docx formatter script");
                Logger.Log("OpenXML Conveer: AddMultipleStyles is done");
                Style.Metatags.StylesCreator.AddMultipleStyles(_doc, new Dictionary<string, string>() {
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
                return false;
            }
            finally
            {
                Logger.Log("OpenXML: file formatted.");
            }
            return true;
        }
        public bool Save()
        {
            try
            {
                _doc.Save();
            }
            catch (Exception ex)
            {
                Logger.Log($"Saving - {ex.ToString()}", LoggerState.Error);
                return false;
            }
            finally
            {
                Logger.Log("OpenXML: file saved.");
            }
            return true;
        }
    }
}