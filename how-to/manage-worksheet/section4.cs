using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ManageWorksheet
{
    public static class Section4
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");
            
            // Remove workSheet1
            workBook.RemoveWorkSheet(1);
            
            // Remove workSheet2
            workBook.RemoveWorkSheet("workSheet2");
            
            workBook.SaveAs("removeWorksheet.xlsx");
        }
    }
}