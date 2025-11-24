using core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace DocXFunc.Style.@struct
{
    public static class BaseText
    {

        public static void allTextStyle(DocX doc)
        {
            doc.MarginLeft = 0;    // 0 см слева
            doc.MarginRight = 0;    // 0 см справа
            doc.MarginTop = 0;    // 0 см сверху (или 720f = 2.54 см, если нужно)
            doc.MarginBottom = 0;
            foreach (var item in doc.Paragraphs)
            {
                item.LineSpacing = 12f;
                item.StyleName = null;
                item.ClearBookmarks();
            }

            //paragraph.IndentationBefore = 0;
            //paragraph.IndentationAfter = 0;
            //paragraph.IndentationHanging = 0;

        }



        public static void BaseTextStyle(Paragraph paragraph)
        {
            string text = paragraph.Text;

            // 2. Очищаем параграф (удаляем всё содержимое)
            paragraph.Xml.RemoveAll();

            // 3. Вставляем текст заново — теперь он "чистый", без стиля
            paragraph.Append(text);
            //paragraph.Bold(false);
            paragraph.Font("Times New Roman");
            paragraph.FontSize(Constants.MainFontSize);
            paragraph.Alignment = Alignment.both;
            paragraph.IndentationFirstLine = 35.43f;

        }




    }
}
