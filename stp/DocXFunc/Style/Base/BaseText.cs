using core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace DocXFunc.Style.Base
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
                item.StyleId = null;
                item.StyleName = null;
                item.ClearBookmarks();
            }
            //paragraph.IndentationBefore = 0;
            //paragraph.IndentationAfter = 0;
            //paragraph.IndentationHanging = 0;
        }

        public static void HeaderOneLevel(Paragraph paragraph)
        {
            paragraph.IndentationFirstLine = 0;
            paragraph.Alignment = Alignment.center;
            paragraph.Bold(true);
            paragraph.FontSize(16);
        }


        public static void BaseTextStyle(Paragraph paragraph)
        {
            paragraph.StyleId = null;
            paragraph.StyleName = null;
            //string text = paragraph.Text;
            //paragraph.Xml.RemoveAll();
            //paragraph.Append(text);
            //paragraph.Bold(false);
            paragraph.SpacingAfter(0);
            paragraph.SpacingBefore(0);
            paragraph.Font("Times New Roman");
            paragraph.FontSize(Constants.MainFontSize);
            paragraph.Alignment = Alignment.both;
            paragraph.IndentationFirstLine = Constants.FirstLineIndent;

        }




    }
}
