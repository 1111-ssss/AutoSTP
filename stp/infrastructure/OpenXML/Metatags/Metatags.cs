using DocumentFormat.OpenXml.Packaging;
using infrastructure.Context;

namespace infrastructure.OpenXML.Metatags
{
    public class Metatags
    {
        public static void AddMetatags(OpenXmlContext context, string author = "", string title = "", string subject = "", string description = "", string lastModifiedBy = "")
        {
            context.Doc.PackageProperties.Creator = author;
            context.Doc.PackageProperties.Title = title;
            context.Doc.PackageProperties.Subject = subject;
            context.Doc.PackageProperties.Description = description;
            context.Doc.PackageProperties.LastModifiedBy = lastModifiedBy;
        }
    }
}