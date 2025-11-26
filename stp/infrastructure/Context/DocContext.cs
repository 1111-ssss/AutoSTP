using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace infrastructure.OpenXmlContext
{
    public class DocContext
    {
        public WordprocessingDocument Doc { get; private set; }

        public MainDocumentPart MainDocumentPart { get; private set; }
        
        public Body Body { get; private set; }

        public DocContext(WordprocessingDocument doc) 
        {
            Doc = doc;

            MainDocumentPart = doc.MainDocumentPart ?? doc.AddMainDocumentPart();
            if (MainDocumentPart.Document.Body == null)
            {
                MainDocumentPart.Document.Body = new Body();
            }
            Body = MainDocumentPart.Document.Body;

         


        }
    }
}
