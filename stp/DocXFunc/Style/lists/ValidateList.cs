using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;
using core;
namespace DocXFunc.Style.links
{
    public static class ValidateList
    {
        public static void Validate(Paragraph paragraph)
        {
            


            if (paragraph.ListItemType == ListItemType.Bulleted)
            {
                Console.WriteLine("Маркированный список");
            }
            else if (paragraph.ListItemType == ListItemType.Numbered)
            {
                Console.WriteLine("Нумерованный список");
            }
            else
            {
               Console.WriteLine("список неподоходящего типа");
            }
        }


        public static void ProcessList(List list)  
        {
             
            foreach (var item in list.Items)
            {
                item.IndentationFirstLine = Constants.FirstLineIndent;
                item.Font("Times New Roman").FontSize(14);  // Шрифт по стандарту
                item.LineSpacing = 18f;  // 1.5 интервал
            }
        }



    }
}
