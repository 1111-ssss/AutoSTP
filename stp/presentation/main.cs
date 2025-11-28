using System.Globalization;
using core.Model;
using infrastructure.Validate.Utils.ArgParser;
using infrastructure.Validate;
using infrastructure.Context;
using core.Enums;
using application.Pipelines;

class Program
{
    static void Main(String[] args)
    {
        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;

        AppOptions options = ArgParser.Parse(args);
        bool valid = ValidateArgs.ValidateOptions(options);

        if (!valid) return;

        var docxContext = new DocXContext(options.InputFile);
        var docxPipeline = PipelineFactory.CreateDocXPipeline(docxContext, options.FileType);
        docxPipeline.StartPipeline();

        if (options.Save)
        {
            options.OutputPath = options.InputFile;
            docxPipeline.Save();
        }
        else
        {
            docxPipeline.SaveAS(options.OutputPath);
        }

        var openXmlContext = new OpenXmlContext(options.InputFile);
        var openXmlPipeline = PipelineFactory.CreateOpenXMLPipeline(openXmlContext, options.FileType);
        openXmlPipeline.StartPipeline();

        openXmlPipeline.Save();

        ValidateArgs.RenameFile(ref options);
        ValidateArgs.OpenFile(options.OutputPath);
    }
}