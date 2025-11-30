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


namespace application.Pipelines.DocX
{
    public class PracticalWorkDocXPipeline : IPipeline
    {
        private readonly DocXContext _context;
        public PracticalWorkDocXPipeline(DocXContext doc, PracticalWorkOptions options)
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
