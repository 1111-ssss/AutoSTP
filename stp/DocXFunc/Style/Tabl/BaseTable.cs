using core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;
using DocXFunc.ex;

namespace DocXFunc.Style.Tabl
{
    public static class BaseTable
    {

        public static void BaseTableStyle(Table table)
        {
            
            var fontSize = Constants.TableFontSize;
            table.Alignment = Alignment.left;
            if(table.RowCount < 20)
            {
                fontSize = Constants.MainFontSize;
            }
            foreach(var item in table.Paragraphs)
            {
                item.FontSize(fontSize);
                item.IndentationFirstLine = 0;
                item.Alignment = Alignment.left;    
            }
            ValidateTable(table, fontSize);
        }


        private static void ValidateTable(Table table, float fontSize)
        {
            foreach(var item in table.Rows[0].Paragraphs)
            {
                item.Alignment = Alignment.center;
            }

            foreach (var row in table.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    foreach (var p in cell.Paragraphs)
                    {
        
                            p.FontSize(fontSize);
            
                    }
                    //SetCellVerticalAlignmentTop(cell);
                }
                
            }

            
        }


        private static void SetCellVerticalAlignmentTop(Cell cell)
        {
            // Получаем TableCellProperties
            var tcp = cell.Xml.Element(XName.Get("tcPr", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"))
                      ?? cell.Xml.AddElement("tcPr");

            // Ищем или создаём vAlign
            var vAlign = tcp.Element(XName.Get("vAlign", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));
            if (vAlign == null)
            {
                vAlign = new XElement(XName.Get("vAlign", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));
                tcp.Add(vAlign);
            }

            // Устанавливаем значение "top"
            vAlign.SetAttributeValue(XName.Get("val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main"), "top");
        }


    }
}
