using DocumentFormat.OpenXml.Packaging;
using WP = DocumentFormat.OpenXml.Wordprocessing;

namespace infrastructure.OpenXML.Metatags
{
    public class StylesCreator
    {
        public static void AddMultipleStyles(MainDocumentPart mainPart, Dictionary<string, string> styles)
        {
            var stylesPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
            if (stylesPart.Styles == null) stylesPart.Styles = new WP.Styles();

            foreach (var style in styles)
            {
                stylesPart.Styles.Append(CreateStyle(style.Key, style.Value));
            }
        }
        private static WP.Style CreateStyle(string styleId, string displayName)
        {
            var style = new WP.Style
            {
                Type = WP.StyleValues.Paragraph,
                StyleId = styleId
            };
            style.StyleName = new WP.StyleName { Val = displayName };
            style.PrimaryStyle = new WP.PrimaryStyle();

            return style;
        }
    }
}