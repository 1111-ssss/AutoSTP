using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Xceed.Document.NET;
//using 

namespace infrastructure.OpenXML.Style
{
    public static class TableStyle
    {
        public static void ApplyBaseTableStyle(Body body)
        {
            var tables = body.Descendants<Table>().ToList();
            foreach (var table in body.Descendants<Table>())
            {
                var tblPr = table.GetFirstChild<TableProperties>();
                if (tblPr == null)
                {
                    tblPr = new TableProperties();
                    table.PrependChild(tblPr);
                }

                // Правильный способ — через object initializer (да, с "серым", но это норма!)
                tblPr.TableStyle = new DocumentFormat.OpenXml.Wordprocessing.TableStyle() { Val = "TableGrid" };
                

                //tblPr.TableWidth = new TableWidth
                //{
                //    Width = "0",
                //    Type = TableWidthUnitValues.Pct
                //};

                //tblPr.TableCellMarginDefault = new TableCellMarginDefault
                //{
                //    //LeftMargin = new LeftMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                //    //RightMargin = new RightMargin { Width = "100", Type = TableWidthUnitValues.Dxa },
                //    TopMargin = new TopMargin { Width = "80", Type = TableWidthUnitValues.Dxa },
                //    BottomMargin = new BottomMargin { Width = "80", Type = TableWidthUnitValues.Dxa }
                //};
            }
        }
        }
    }

