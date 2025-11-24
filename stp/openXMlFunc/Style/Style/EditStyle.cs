using openXMlFunc.Converter;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace openXMlFunc.EditStyle
{
    class EditStyle
    {
        public static void ApplyBaseStyle(WordprocessingDocument doc)
        {
            var mainPart = doc.MainDocumentPart;
            if (mainPart == null) return;

            var stylesPart = mainPart.StyleDefinitionsPart;
            if (stylesPart == null)
            {
                stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
                var newStyles = new Styles();
                CreateOrUpdateNormalStyle(newStyles);
                stylesPart.Styles = newStyles;
                return;
            }

            // Если StylesPart уже существует — работаем с существующим Styles
            var styles = stylesPart.Styles;
            if (styles == null)
            {
                styles = new Styles();
                stylesPart.Styles = styles;
            }

            // Обновляем стиль в существующем объекте
            CreateOrUpdateNormalStyle(styles);
        }
        private static void CreateOrUpdateNormalStyle(Styles styles)
        {
            var normalStyle = styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == "Normal");
            if (normalStyle == null)
            {
                normalStyle = new Style { Type = StyleValues.Paragraph, StyleId = "Normal" };
                normalStyle.Append(new Name { Val = "Normal" });
                normalStyle.Append(new BasedOn { Val = "Normal" });
                styles.Append(normalStyle);
            }

            normalStyle.StyleRunProperties = new StyleRunProperties(
                new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                new FontSize { Val = UnitConverter.PtToHalfPoints(14) },
                new Color { Val = "000000" }
            );

            var paraProps = normalStyle.StyleParagraphProperties ?? new StyleParagraphProperties();
            paraProps.SpacingBetweenLines = new SpacingBetweenLines
            {
                After = "0",
                Before = "0",
                Line = "240",
                LineRule = LineSpacingRuleValues.Auto
            };
            paraProps.Indentation = new Indentation { FirstLine = UnitConverter.CmToTwips(1.25) };
            paraProps.Justification = new Justification { Val = JustificationValues.Both };

            normalStyle.StyleParagraphProperties = paraProps;
        }
    }
}