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

            }
        }


        public static void BaseTextStyle(Paragraph paragraph)
        {
            paragraph.Font("Times New Roman");
            paragraph.FontSize(Constants.MainFontSize);
            paragraph.Alignment = Alignment.both;
            paragraph.IndentationFirstLine = 35.43f;
        }




    }
}
