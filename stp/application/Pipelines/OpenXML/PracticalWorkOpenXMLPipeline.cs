using core.Enums;
using infrastructure.Context;
using infrastructure.DocX.Lists;
using infrastructure.DocX.MainText;
using infrastructure.DocX.Pictures;
using infrastructure.DocX.Tables;
using infrastructure.Utils.UtilsDocX.Validate;
using logger.Logger;
using core.Interfaces;
using System.Text.RegularExpressions;


namespace application.Pipelines.OpenXML
{
    public class PracticalWorkOpenXMLPipeline : IPipeline
    {
        private readonly OpenXmlContext _context;
        public PracticalWorkOpenXMLPipeline(OpenXmlContext doc)
        {
            _context = doc;
        }
        public void StartPipeline()
        {
            
        }
        public void SaveAS(string TargetPath)
        {

        }
        public void Save()
        {

        }
    }
}
