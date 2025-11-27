using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;
using core.Const;
using logger.Logger;
namespace DocXFunc.Style.links
{
    public static class ValidateList
    {
        public static void Validate(Paragraph paragraph)
        {            

            if (paragraph.ListItemType == ListItemType.Bulleted)
            {
                Logger.Debug("Маркированный список");
            }
            else if (paragraph.ListItemType == ListItemType.Numbered)
            {
                Logger.Debug("Нумерованный список");
            }
            else
            {
               Logger.Debug("список неподоходящего типа");
            }
        }

        public static void ProcessList(Paragraph paragraph)  
        {
            
            paragraph.IndentationFirstLine = Constants.FirstLineIndent;
            paragraph.Font("Times New Roman").FontSize(14);  
            //paragraph.LineSpacing = 18f;  // 1.5 интервал
        }

    }
}
