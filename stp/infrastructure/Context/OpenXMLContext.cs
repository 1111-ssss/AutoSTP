using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using logger.Logger;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace infrastructure.Context
{
    public class OpenXmlContext : IDisposable
    {
        public WordprocessingDocument Doc { get; private set; }

        public MainDocumentPart MainDocumentPart { get; private set; }
        
        public Body Body { get; private set; }

        public OpenXmlContext(String pathtoFile) 
        {
            var doc = WordprocessingDocument.Open(pathtoFile, isEditable: true);

            Doc = doc;
            MainDocumentPart = doc.MainDocumentPart ?? doc.AddMainDocumentPart();
            if (MainDocumentPart.Document.Body == null)
            {
                MainDocumentPart.Document.Body = new Body();
            }
            Body = MainDocumentPart.Document.Body;
        }
        public void Dispose()
        {
            Doc?.Dispose();
        }
    }
}
