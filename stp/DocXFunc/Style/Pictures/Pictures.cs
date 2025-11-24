using core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace DocXFunc.Style.Pictures
{
    public class Pictures
    {
        public static void validate(Paragraph paragraph)
        {
            paragraph.IndentationFirstLine = 0;
            paragraph.Alignment = Alignment.center;
            paragraph.SpacingBefore(6);
            IndeNull(paragraph);
        }
        
        public static void IndeNull(Paragraph paragraph)
        {
            paragraph.IndentationBefore = 0;
            paragraph.IndentationAfter = 0;
            paragraph.IndentationHanging = 0;
        }


        public static void PictureNameStyle(Paragraph paragraph)
        {
            paragraph.Alignment = Alignment.center;
            paragraph.IndentationAfter = Constants.MainFontSize;
            paragraph.Font("Times New Roman");
            paragraph.FontSize(Constants.MainFontSize);
            IndeNull(paragraph);
        }

    }
}
