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
using core.Model;


namespace application.Pipelines.OpenXML
{
    public class GraduateWorkOpenXMLPipeline : IPipeline
    {
        private readonly OpenXmlContext _context;
        public GraduateWorkOpenXMLPipeline(OpenXmlContext doc, GraduateWorkOptions options)
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
