using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace openXMlFunc.Style.Margins
{
    class Margins
    {
        public static void SetupPageMargins(WordprocessingDocument doc)
        {
            var mainPart = doc.MainDocumentPart ?? doc.AddMainDocumentPart();
            if (mainPart.Document.Body == null)
            {
                mainPart.Document.Body = new Body();
            }
            var body = mainPart.Document.Body;

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