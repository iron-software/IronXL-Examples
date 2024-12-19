using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
        }
    }
}