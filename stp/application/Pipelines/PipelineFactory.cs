using core.Enums;
using core.Interfaces;
using infrastructure.Context;
using application.Pipelines.DocX;
using application.Pipelines.OpenXML;
using core.Model;

namespace application.Pipelines
{
    public class PipelineFactory
    {
        private readonly FileType _workType;
        public static IPipeline CreateDocXPipeline(DocXContext context, AppOptions options)
        {
            return options.FileType switch
            {
                FileType.LabWork => new LabWorkDocXPipeline(context, (LabWorkOptions)options),
                FileType.PracticalWork => new PracticalWorkDocXPipeline(context, (PracticalWorkOptions)options),
                FileType.GraduateWork => new GraduateWorkDocXPipeline(context, (GraduateWorkOptions)options),
                _ => throw new ArgumentException("Неизвестный тип работы")
            };
        }

        public static IPipeline CreateOpenXMLPipeline(OpenXmlContext context, AppOptions options)
        {
            return options.FileType switch
            {
                FileType.LabWork => new LabWorkOpenXMLPipeline(context, (LabWorkOptions)options),
                FileType.PracticalWork => new PracticalWorkOpenXMLPipeline(context, (PracticalWorkOptions)options),
                FileType.GraduateWork => new GraduateWorkOpenXMLPipeline(context, (GraduateWorkOptions)options),
                _ => throw new ArgumentException("Неизвестный тип работы")
            };
        }
    }
}