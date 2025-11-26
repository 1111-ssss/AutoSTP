using infrastructure.DocX.Style.MainText;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;
using dc = Xceed.Document.NET;

namespace infrastructure.Context
{
    public class DocXContext
    {
        public Xceed.Words.NET.DocX Doc { get; private set; }
        public DocXContext(string SoursePath) 
        {
            Doc = Xceed.Words.NET.DocX.Load(SoursePath);
            MainTextStyle.allTextStyle(Doc);
        
        }
    }
}
