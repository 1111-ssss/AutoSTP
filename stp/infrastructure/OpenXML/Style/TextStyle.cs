using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;

namespace infrastructure.OpenXML.Style
{
    public static class TextStyle
    {
        public static void ApplyBaseTextStyle(Body body)
        {
            foreach (var run in body.Descendants<Run>())
            {
                RunProperties pr = run.RunProperties ?? run.PrependChild(new RunProperties());

                pr.RunFonts = new RunFonts()
                {
                    Ascii = "Times New Roman",
                    HighAnsi = "Times New Roman",
                    ComplexScript = "Times New Roman"
                };    
                

            }

        }


    }
}
