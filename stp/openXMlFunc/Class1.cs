// using DocumentFormat.OpenXml;
// using DocumentFormat.OpenXml.Packaging;
// using DocumentFormat.OpenXml.Wordprocessing;
// using System;
// using System.Drawing;
// using System.IO;


// namespace openXMlFunc
// {

//     public static class class1
//     {
 


// public static void ReplaceStylesWithFourCustom(string docxPath)
//     {
//         using (var doc = WordprocessingDocument.Open(docxPath, true))
//         {
//             // Удаляем старый styles.xml
//             var stylesPart = doc.MainDocumentPart.StyleDefinitionsPart;
//             if (stylesPart != null)
//             {
//                 doc.MainDocumentPart.DeletePart(stylesPart);
//             }

//             // Создаём новый
//             stylesPart = doc.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

//             // Обязательно: корневой элемент + namespace
//             var styles = new Styles();
//             styles.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

//             // === Стиль 1: СДЕЛАНО ===
//             var style1 = CreateStyle("STP_SDELANO", "СДЕЛАНО");

//             // === Стиль 2: ГУРСКИЙ ===
//             var style2 = CreateStyle("STP_GURSKIY", "ГУРСКИЙ");

//             // === Стиль 3: И ===
//             var style3 = CreateStyle("STP_I", "И");

//             // === Стиль 4: КАСЕВИЧ ===
//             var style4 = CreateStyle("STP_KASEVICH", "КАСЕВИЧ");

//             // Добавляем в нужном порядке
//             styles.Append(style1);
//             styles.Append(style2);
//             styles.Append(style3);
//             styles.Append(style4);

//             // Сохраняем
//             stylesPart.Styles = styles;
//         }
//     }


//     private static Style CreateStyle(string styleId, string displayName)
//     {
//         var style = new Style
//         {
//             Type = StyleValues.Paragraph,
//             StyleId = styleId
//         };
//         style.StyleName = new StyleName { Val = displayName };
//         style.PrimaryStyle = new PrimaryStyle();

//         // Обязательно: добавляем хотя бы базовое форматирование
//         var runProps = new StyleRunProperties();
//         runProps.Append(new RunFonts { Ascii = "Times New Roman" });
//         runProps.Append(new FontSize { Val = $"40" }); // 14 pt

//         var paraProps = new StyleParagraphProperties();
//         paraProps.Append(new SpacingBetweenLines { Line = "240" }); // одинарный интервал
//         paraProps.Append(new Indentation { FirstLine = "709" });     // 1.25 см
//         paraProps.Append(new Justification { Val = JustificationValues.Both });

//         style.Append(runProps);
//         style.Append(paraProps);

//         return style;
//     }
// }
// }
