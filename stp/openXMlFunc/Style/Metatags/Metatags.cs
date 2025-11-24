using DocumentFormat.OpenXml.Packaging;

namespace openXMlFunc.Metatags
{
    class Metatags
    {
        public static void AddMetatags(WordprocessingDocument doc, String author = "", String title = "", String subject = "", String description = "")
        {
            doc.PackageProperties.Creator = author;
            doc.PackageProperties.Title = title;
            doc.PackageProperties.Subject = subject;
            doc.PackageProperties.Description = description;
        }
    }
}