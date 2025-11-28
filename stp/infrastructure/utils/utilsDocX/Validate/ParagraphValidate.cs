using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace infrastructure.Utils.UtilsDocX.Validate
{
    public static class ParagraphValidate
    {
        public static bool IsSpecialParagraph(Paragraph p)
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
                text.StartsWith("СОДЕРЖАНИЕ") ||
                text.Contains("Лабораторная работа") ||
                text.Contains("Практическая работа") ||
                text.Contains("Вариант работа") ||
                  text.Contains("ВВЕДЕНИЕ") ||
                string.IsNullOrWhiteSpace(text))
                return true;

            // 2. Подписи к рисункам и таблицам
            if (Regex.IsMatch(text, @"^Рисун(ок|.\s*)\s*\d+", RegexOptions.IgnoreCase) ||
                (Regex.IsMatch(text.Trim(), @"^Рисунок \d+\.\d+ –", RegexOptions.IgnoreCase)
                ))
                return true;




            if (text == "СОДЕРЖАНИЕ" ||
                text == "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ" ||
                text.Contains("ПРИЛОЖЕНИЕ") ||
                Regex.IsMatch(text, @"^\d+\s+[А-ЯA-Z]"))
                return true;


            if (p.Pictures.Any() == false && p.Pictures.Count > 0)
                return true;

            if (string.IsNullOrWhiteSpace(text) || text.Length < 3)
                return true;

            return false;
        }
    }
}

