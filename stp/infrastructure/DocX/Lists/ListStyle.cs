using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;
using core.Const;
namespace infrastructure.DocX.Lists
{
    public static class ListStyle
    {
        public static void ProcessList(Paragraph paragraph)  
        {
            paragraph.IndentationFirstLine = Constants.FirstLineIndent;
            paragraph.Font("Times New Roman").FontSize(14);  
            //paragraph.LineSpacing = 18f;  // 1.5 интервал
        }

    }
}
