
using System.Runtime.InteropServices;
using System.IO;
using DocXFunc;


class Program
{
    static void Main()
    {
        // 1. Запускаем Word (в фоне)
        DoXcConveer doXcConveer = new DoXcConveer(@"C:\Users\Rog\Desktop\ТПО ПР1.docx");
        doXcConveer.AllConveer();
        doXcConveer.Save(@"C:\Users\Rog\Desktop\ТПО ЛР22 Касевич В.А..docx");
    }
}