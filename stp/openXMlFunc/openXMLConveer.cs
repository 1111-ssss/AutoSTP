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
                Logger.Debug("OpenXML Conveer: SetupPageMargins is done");
                // Style.Style.EditStyle.CreateBaseStyle(_doc);
                // logger.Logger.Log("OpenXML Conveer: CreateBaseStyle is done");
                Style.Style.TextStyle.ApplyBaseStyle(_doc);
                Logger.Debug("OpenXML Conveer: ApplyBaseStyle is done");
                Style.Names.TableNames.TableName(_doc, 1);
                Logger.Debug("OpenXML Conveer: TableNames is done");
                Style.Metatags.Metatags.AddMetatags(_doc, author: "AutoSTP", description: "Document formatted with private docx formatter script");
                Logger.Debug("OpenXML Conveer: AddMultipleStyles is done");
                Style.Metatags.StylesCreator.AddMultipleStyles(_doc, new Dictionary<string, string>() {
                    {"stp1", "1 Automatical"},
                    {"stp2", "2 STP"},
                    {"stp3", "3 Formatter"},
                    {"stp4", "4 Script"},
                });
                Logger.Debug("OpenXML Conveer: AddMetatags is done");
            }
            catch (Exception ex)
            {
                Logger.Error($"Formatting - {ex.ToString()}");
                return false;
            }
            finally
            {
                Logger.Debug("OpenXML: file formatted.");
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
                Logger.Error($"Saving - {ex.ToString()}");
                return false;
            }
            finally
            {
                Logger.Debug("OpenXML: file saved.");
            }
            return true;
        }
    }
}