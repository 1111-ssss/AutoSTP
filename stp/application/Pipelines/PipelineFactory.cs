using core.Enums;
using core.Interfaces;
using infrastructure.Context;
using application.Pipelines.DocX;
using application.Pipelines.OpenXML;

namespace application.Pipelines
{
    public class PipelineFactory
    {
        private readonly FileType _workType;
        public static IPipeline CreateDocXPipeline(DocXContext context, FileType workType)
        {
            return workType switch
            {
                FileType.LabWork => new LabWorkDocXPipeline(context),
                FileType.PracticalWork => new PracticalWorkDocXPipeline(context),
                FileType.GraduateWork => new GraduateWorkDocXPipeline(context),
                // _ => throw new ArgumentException("Неизвестный тип работы")
                _ => new LabWorkDocXPipeline(context),
            };
        }

        public static IPipeline CreateOpenXMLPipeline(OpenXmlContext context, FileType workType)
        {
            return workType switch
            {
                FileType.LabWork => new LabWorkOpenXMLPipeline(context),
                FileType.PracticalWork => new PracticalWorkOpenXMLPipeline(context),
                FileType.GraduateWork => new GraduateWorkOpenXMLPipeline(context),
                // _ => throw new ArgumentException("Неизвестный тип работы")
                _ => new LabWorkOpenXMLPipeline(context),
            };
        }
    }
}