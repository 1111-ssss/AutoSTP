using core.Const;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Xceed.Document.NET;
using Xceed.Words.NET;


namespace infrastructure.DocX.Tables
{
    public static class MainTable
    {
        public static void TableNameStyle(Paragraph paragraph)
        {
            paragraph.SpacingBefore(6);
            paragraph.IndentationFirstLine = 0;
            paragraph.Alignment = Alignment.left;
            paragraph.Font("Times New Roman").FontSize(14);

        }


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


    }
}
