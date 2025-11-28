using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using logger.Logger;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace infrastructure.OpenXML.NamesProcessing
{
    public class PicturesNames
    {
        private Body _body;

        private int _chapter;

        public PicturesNames(Body body, int chapter = 1)
        {
            _body = body;
            _chapter = chapter;
        }


        public void AddPictureNames()
        {
            AddNamesConveer<Picture>();
            AddNamesConveer<Drawing>();
        }


        public void AddNamesConveer<T>() where T : OpenXmlElement
        {
          

            var pictures = _body.Descendants<T>().ToList();
            int count = 0;
            Logger.Debug($"количество рисунков {pictures.Count}");


            foreach (var picture in pictures)
            {
                count++;
                var caption = GetParagraphAfter(picture.Ancestors<Paragraph>().FirstOrDefault());
                if (caption == null)
                {
                    Logger.Debug("caption is null");
                }
                if (caption == null || !Regex.IsMatch(caption.InnerText.Trim(), @"^Рисунок \d+\.\d+ –", RegexOptions.IgnoreCase))
                {
                    if (caption != null)
                    {
                        Logger.Info($"текст после рисунка {caption.InnerText.Trim()}");
                    }



                    Paragraph paragraph = CreateParagrapg(
                        JustificationValues.Center,
                        $"Рисунок {_chapter}.{count} – ",
                        new Indentation() { 
                            FirstLine = 0.ToString()
                        
                        }

                    );

                    Logger.Debug("подпись рисунка не найдена");

                    Logger.Debug($"adding pargraph {paragraph.InnerText}");

                    var outt = picture.Ancestors<Paragraph>().FirstOrDefault().InsertAfterSelf(paragraph);

                    Logger.Debug($"added {outt.InnerText}");


                    //

                }




            }

        }


        private Paragraph CreateParagrapg
            (
            JustificationValues justificationvalue,
            string Text,
            Indentation indentation,
            string Font = "Times New Roman",
            int FontSize = 14
            )
        {
            var paragraph = new Paragraph();

            var props = new ParagraphProperties();

            JustificationValues justification = justificationvalue;
            props.Justification = new Justification() { Val = justification };
            props.Indentation = indentation;

            paragraph.Append(props);

            var run = new Run();
            var runProps = new RunProperties();

            runProps.FontSize = new FontSize() { Val = (FontSize * 2).ToString() };
            
            var font = "Times New Roman";
            runProps.RunFonts = new RunFonts()
            {
                Ascii = font,
                HighAnsi = font,
                ComplexScript = font
            };



            var pPr = paragraph.ParagraphProperties ?? new ParagraphProperties();
            pPr.SpacingBetweenLines = new SpacingBetweenLines()
            {
                Before = 0.ToString(),
                After = 120.ToString(),
                Line = "240",
                LineRule = LineSpacingRuleValues.Auto
                
            };

            paragraph.ParagraphProperties = pPr;

            Text text = new Text(Text);
            run.Append(runProps);
            run.Append(text);
            paragraph.Append(run);
            


            return paragraph;

           
          
        }


        private Paragraph? GetParagraphAfter(Paragraph picture)
        {

            var next = picture.NextSibling<Paragraph>();
            while (next != null)
            {
                // Пропускаем пустые абзацы и разрывы страниц
                if (!string.IsNullOrWhiteSpace(next.InnerText))
                    return next;

                next = next.NextSibling<Paragraph>();
            }
            return null;
        }


    }
}

