using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ClearCells
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            
            // Delete all worksheets
            workBook.WorkSheets.Clear();
            
            workBook.SaveAs("useClear.xlsx");
        }
    }
}