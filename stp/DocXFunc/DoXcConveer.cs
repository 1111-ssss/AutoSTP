using DocXFunc.Style.Pictures;
using DocXFunc.Style.@struct;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;
using static System.Net.Mime.MediaTypeNames;
using core;

namespace DocXFunc
{
    public class DoXcConveer
    {
        private readonly DocX doc;
        public DoXcConveer(string Soursepath)
        {
            doc = DocX.Load(Soursepath);
            BaseText.allTextStyle(doc);
        }

        public void ParagraphConveer()
        {
            foreach (var item in doc.Paragraphs)
            {
                if (IsSpecialParagraph(item)) continue;                
                    
                BaseText.BaseTextStyle(item);
                
            }
        }

        public void PicturesConveer()
        {
          foreach(var pr in doc.Paragraphs)
           {
                if(pr.Pictures.Any())
                {
                    Pictures.validate(pr);
                }
           }
        }

        public void TableConveer()
        {
            foreach(var item in doc.Tables)
            {
                BaseTable.BaseTableStyle(item);
            }
        }


        public void AllConveer()
        {
            foreach( var item in doc.Paragraphs)
            {
                if (!IsSpecialParagraph(item))
                {
                    BaseText.BaseTextStyle(item);
                }
                
                if(item.Pictures.Any())
                {
                    Pictures.validate(item);
                }

                if(Regex.IsMatch(item.Text.Trim(), @"^Рисун(ок|.\s*)\s*\d+", RegexOptions.IgnoreCase))
                {
                    item.Alignment = Alignment.center;
                    item.IndentationAfter = Constants.MainFontSize;
                    item.Font("Times New Roman");
                    item.FontSize(Constants.MainFontSize);
                }
               
            }
            TableConveer();
        }


        public void Save(string TargetPath)
        {
            doc.SaveAs(TargetPath);
        }

        private bool IsSpecialParagraph(Paragraph p)
        {
            string text = p.Text.Trim();

            // 1. Титульный лист — пропускаем
            if (
                //text.Contains("Полоцкий государственный экономический колледж") ||
                //text.Contains("СТП ПГЭК") ||
                //text.Contains("Учреждение образования") ||
                //text.Contains("Полоцк 2023") ||
                //text.Contains("Рассмотрен и одобрен") ||
                //text.Contains("Утвержден приказом") ||
                text.Contains("Лабораторная работа") ||
                text.Contains("Практическая работа") ||
                text.Contains("Вариант работа") ||
                string.IsNullOrWhiteSpace(text))
                return true;

            // 2. Подписи к рисункам и таблицам
            if (Regex.IsMatch(text, @"^Рисун(ок|.\s*)\s*\d+", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(text, @"^Табл(ица|.\s*)\s*\d+", RegexOptions.IgnoreCase))
                return true;

            // 3. Заголовки разделов (СОДЕРЖАНИЕ, СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ и т.д.)
            if (text == "СОДЕРЖАНИЕ" ||
                text == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" ||
                text.Contains("ПРИЛОЖЕНИЕ") ||
                Regex.IsMatch(text, @"^\d+\s+[А-ЯA-Z]"))
                return true;

            // 4. Параграфы с картинками (логотип, иллюстрации)
            if (p.Pictures.Any() || p.ParentContainer is Table == false && p.Pictures.Count > 0)
                return true;

            // 5. Пустые строки и разрывы страниц
            if (string.IsNullOrWhiteSpace(text) || text.Length < 3)
                return true;

            return false;
        }
    }
}
