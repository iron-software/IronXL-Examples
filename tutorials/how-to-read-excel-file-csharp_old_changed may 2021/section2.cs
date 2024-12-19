using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp_Old_Changed May 2021
{
    public static class Section2
    {
        public static void Run()
        {
            /**
            Load WorkBook
            anchor-load-a-workbook
            **/
            var workbook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
        }
    }
}