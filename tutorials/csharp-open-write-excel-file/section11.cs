using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section11
    {
        public static void Run()
        {
            workBook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
        }
    }
}