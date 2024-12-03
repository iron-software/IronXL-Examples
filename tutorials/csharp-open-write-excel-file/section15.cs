using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section15
    {
        public void Run()
        {
            workBook.SaveAsXml($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldXML.XML");
        }
    }
}