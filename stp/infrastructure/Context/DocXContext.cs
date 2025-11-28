using infrastructure.DocX.MainText;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
using dc = Xceed.Document.NET;

namespace infrastructure.Context
{
    public class DocXContext : IDisposable
    {
        public Xceed.Words.NET.DocX Doc { get; private set; }
        public DocXContext(String pathToFile) 
        {

            Doc = Xceed.Words.NET.DocX.Load(pathToFile);
            MainTextStyle.allTextStyle(Doc);
            
        
        }
        public void Dispose()
        {
            Doc?.Dispose();
        }
    }
}
