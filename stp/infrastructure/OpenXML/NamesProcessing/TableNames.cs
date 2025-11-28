using core;
//using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using logger.Logger;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using Justification = DocumentFormat.OpenXml.Wordprocessing.Justification;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace infrastructure.OpenXML.NamesProcessing
{
    public static class TableNames
    {
        public static void TableName(Body body, int chapter = 1)
        {
           
            var tables = body.Descendants<Table>().ToList();
            int count = 0;
            Logger.Debug($"количество таблиц {tables.Count}");
            foreach( var table in tables)
            {
                count++;
                var caption = GetParagraphBeforeTable(table); 
                if(!Regex.IsMatch(caption.InnerText.ToLower(), @"^таблица \d+\.\d+ – ", RegexOptions.IgnoreCase))
                {
                    Logger.Debug("подпись таблицы не найдена");
                    var paragraph = new Paragraph();

                    var props = new ParagraphProperties();

                    JustificationValues justification = JustificationValues.Left;
                    props.Justification = new Justification() { Val = justification };
                    props.Indentation = new Indentation() { 
                        FirstLine = 0.ToString(),
                        Left = 0.ToString(),
                    };

              
                    props.SpacingBetweenLines = new SpacingBetweenLines()
                    {
                        Before = 120.ToString(),
                        After = 0.ToString(),
                        Line = "240",
                        LineRule = LineSpacingRuleValues.Auto
                    };
                    


                    paragraph.Append(props);

                    var run = new Run();
                    var runProps = new RunProperties();

                    runProps.FontSize = new FontSize() { Val = 28.ToString() };

                    var font = "Times New Roman";
                    runProps.RunFonts = new RunFonts()
                    {
                        Ascii = font,
                        HighAnsi = font,
                        ComplexScript = font
                    };

                    run.Append(runProps);
                    run.Append(new Text($"Таблица {chapter}.{count} – "));
                    paragraph.Append(run);
                    Console.WriteLine($"Добавлена подпись таблицы : {paragraph.InnerText}" +
                        $"Отступ у подписи таблицы : {paragraph.ParagraphProperties.Indentation.FirstLine}"
                        );
                    Logger.Debug($"Добавлена подпись таблицы : {paragraph.InnerText}");
                    Logger.Debug($"Отступ у подписи таблицы : {paragraph.ParagraphProperties.Indentation.FirstLine}");
                    table.InsertBeforeSelf(paragraph);


                }
                



            }




        }


        private static Paragraph GetParagraphBeforeTable(Table table)
        {
            
            var previous = table.PreviousSibling<Paragraph>();
            while (previous != null)
            {
                // Пропускаем пустые абзацы и разрывы страниц
                if (!string.IsNullOrWhiteSpace(previous.InnerText))
                    return previous;

                previous = previous.PreviousSibling<Paragraph>();
            }
            return null;
        }


    }
}
