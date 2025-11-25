using core.Const;
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
            doc.MarginLeft = 0;    
            doc.MarginRight = 0;   
            doc.MarginTop = 0;    
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

            paragraph.Xml.RemoveAll();
            paragraph.Append(text);
            //paragraph.Bold(false);
            paragraph.Font("Times New Roman");
            paragraph.FontSize(Constants.MainFontSize);
            paragraph.Alignment = Alignment.both;
            paragraph.IndentationFirstLine = 35.43f;

        }




    }
}
