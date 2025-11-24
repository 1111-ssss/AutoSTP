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
            
        }
        
    }
}
