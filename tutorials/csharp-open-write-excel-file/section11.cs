using IronXL;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile
{
    public static class Section11
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");

            workBook.SaveAsJson($@"{Directory.GetCurrentDirectory()}\Files\HelloWorldJSON.json");
        }
    }
}