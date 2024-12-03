using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section13
    {
        public void Run()
        {
            workBook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
        }
    }
}