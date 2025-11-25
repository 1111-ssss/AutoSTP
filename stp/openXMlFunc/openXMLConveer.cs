using core;
using DocumentFormat.OpenXml.Packaging;
using logger;
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
                openXMlFunc.EditStyle.EditStyle.CreateBaseStyle(_doc);
                logger.Logger.Log("OpenXML Conveer: CreateBaseStyle is done");
                openXMlFunc.Style.TextStyle.ApplyBaseStyle(_doc);
                logger.Logger.Log("OpenXML Conveer: ApplyBaseStyle is done");
                openXMlFunc.Margins.Margins.SetupPageMargins(_doc);
                logger.Logger.Log("OpenXML Conveer: SetupPageMargins is done");
                openXMlFunc.Metatags.Metatags.AddMetatags(_doc, author: "AutoSTP", description: "Document formatted with private docx formatter script");
                logger.Logger.Log("OpenXML Conveer: AddMetatags is done");
            }
            catch (Exception ex)
            {
                logger.Logger.Log($"Форматирование - {ex.ToString()}", loggerState: LoggerState.Error);
                return false;
            }
            finally
            {
                logger.Logger.Log("OpenXML: файл отформатирован.");
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
                logger.Logger.Log($"Сохранение - {ex.ToString()}", loggerState: LoggerState.Error);
                return false;
            }
            finally
            {
                logger.Logger.Log("OpenXML: Файл сохранен.");
            }
            return true;
        }
    }
}