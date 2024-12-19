using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ManageWorksheet
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");
            
            // Set worksheet position
            workBook.SetSheetPosition("workSheet2", 0);
            
            workBook.SaveAs("setWorksheetPosition.xlsx");
        }
    }
}