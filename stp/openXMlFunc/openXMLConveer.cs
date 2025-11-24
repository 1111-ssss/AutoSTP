using core;
using DocumentFormat.OpenXml.Packaging;
using logger;

namespace openXMlFunc
{
    public class OpenXMLConveer
    {
        private WordprocessingDocument _doc;
        private String _stylesString;
        public OpenXMLConveer(WordprocessingDocument doc, String stylesString = "Auto STP Script test by НЕГурский and НЕКасевич")
        {
            _doc = doc;
            _stylesString = stylesString;
        }
        public bool AllConveer()
        {
            try
            {
                openXMlFunc.Margins.Margins.SetupPageMargins(_doc);
                openXMlFunc.EditStyle.EditStyle.ApplyBaseStyle(_doc);
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