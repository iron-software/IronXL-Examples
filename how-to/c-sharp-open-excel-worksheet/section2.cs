using IronXL.Excel;
namespace IronXL.Examples.HowTo.CSharpOpenExcelWorksheet
{
    public static class Section2
    {
        public static void Run()
        {
            WorkSheet ws = WorkBook.GetWorkSheet("SheetName");
        }
    }
}