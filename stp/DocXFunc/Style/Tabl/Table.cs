using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;

namespace DocXFunc.Style.Tabl
{
    public class Tables
    {
        public static void TableNameStyle(Paragraph paragraph)
        {
            paragraph.IndentationFirstLine = 0;
            paragraph.Alignment = Alignment.left;
            paragraph.Font("Times New Roman").FontSize(14);

        }
    }
}
