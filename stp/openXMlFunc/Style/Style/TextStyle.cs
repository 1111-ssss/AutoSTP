using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Text;
using System.Threading.Tasks;

namespace openXMlFunc.Style
{
    public static class TextStyle
    {
        public static void ApplyBaseStyle(WordprocessingDocument doc)
        {

            var mainPart = doc.MainDocumentPart ?? doc.AddMainDocumentPart();
            var body = mainPart.Document.Body;
            if (body == null)
            {
                mainPart.Document.Body = new Body();
                body = mainPart.Document.Body;
            }

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
