using core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
                SetupPageMargins();
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
        private void SetupPageMargins()
        {
            var mainPart = _doc.MainDocumentPart ?? _doc.AddMainDocumentPart();
            var body = mainPart.Document.Body;
            if (body == null)
            {
                mainPart.Document.Body = new Body();
                body = mainPart.Document.Body;
            }

            // Получаем или создаём SectionProperties
            var sectPr = body.Elements<SectionProperties>().LastOrDefault();
            if (sectPr == null)
            {
                sectPr = new SectionProperties();
                body.Append(sectPr);
            }

            // Получаем существующий PageMargin или создаём новый
            var pageMargin = sectPr.GetFirstChild<PageMargin>();
            if (pageMargin == null)
            {
                pageMargin = new PageMargin();
                sectPr.Append(pageMargin);
            }

            // Устанавливаем значения (в twips)
            pageMargin.Left = (UInt32Value)1701U;  // ~30 мм
            pageMargin.Right = (UInt32Value)850U;   // ~15 мм
            pageMargin.Top = (Int32Value)1134;    // ~20 мм
            pageMargin.Bottom = (Int32Value)1134;    // ~20 мм
            pageMargin.Header = (UInt32Value)708U;   // ~12.5 мм — расстояние от верхнего края до колонтитула
            pageMargin.Footer = (UInt32Value)708U;   // ~12.5 мм — расстояние от нижнего края до колонтитула
        }
    }
}