
using System.Runtime.InteropServices;
using System.IO;
using DocXFunc;
using System.Globalization;
using openXMlFunc;

class Program
{
    static void Main()
    {
        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.DefaultThreadCurrentUICulture = CultureInfo.InvariantCulture;


        DoXcConveer doXcConveer = new DoXcConveer(@"C:\Users\Rog\Desktop\Отчет 19 Гурский.docx");
        doXcConveer.AllConveer();
        doXcConveer.Save();
        class1.ReplaceStylesWithFourCustom(@"C:\Users\Rog\Desktop\Отчет 19 Гурский.docx");
    }
}