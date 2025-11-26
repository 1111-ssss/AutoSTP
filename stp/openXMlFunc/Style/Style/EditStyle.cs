using openXMlFunc.Style.Converter;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace openXMlFunc.Style.Style
{
    class EditStyle
    {
        public static void CreateBaseStyle(WordprocessingDocument doc)
        {
            var mainPart = doc.MainDocumentPart;
            if (mainPart == null) return;

            var stylesPart = mainPart.StyleDefinitionsPart;
            if (stylesPart == null)
            {
                stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var newStyles = new WP.Styles();
                CreateOrUpdateNormalStyle(newStyles);
                stylesPart.Styles = newStyles;
                return;
            }

            // Если StylesPart уже существует — работаем с существующим Styles
            var styles = stylesPart.Styles;
            if (styles == null)
            {
                styles = new WP.Styles();
                stylesPart.Styles = styles;
            }

            // Обновляем стиль в существующем объекте
            CreateOrUpdateNormalStyle(styles);
        }
        private static void CreateOrUpdateNormalStyle(WP.Styles styles)
        {
            var normalStyle = styles.Elements<WP.Style>().FirstOrDefault(s => s.StyleId?.Value == "Normal");
            if (normalStyle == null)
            {
                normalStyle = new WP.Style { Type = WP.StyleValues.Paragraph, StyleId = "Normal" };
                normalStyle.Append(new WP.Name { Val = "Normal" });
                normalStyle.Append(new WP.BasedOn { Val = "Normal" });
                styles.Append(normalStyle);
            }

            normalStyle.StyleRunProperties = new WP.StyleRunProperties(
                new WP.RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                new WP.FontSize { Val = UnitConverter.PtToHalfPoints(14) },
                new WP.Color { Val = "000000" }
            );

            var paraProps = normalStyle.StyleParagraphProperties ?? new WP.StyleParagraphProperties();
            paraProps.SpacingBetweenLines = new WP.SpacingBetweenLines
            {
                After = "0",
                Before = "0",
                Line = "240",
                LineRule = WP.LineSpacingRuleValues.Auto
            };
            paraProps.Indentation = new WP.Indentation { FirstLine = UnitConverter.CmToTwips(1.25) };
            paraProps.Justification = new WP.Justification { Val = WP.JustificationValues.Both };

            normalStyle.StyleParagraphProperties = paraProps;
        }
    }
}