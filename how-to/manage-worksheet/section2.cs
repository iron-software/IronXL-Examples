using IronXL;
using IronXL.Excel;
namespace ironxl.ManageWorksheet
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");
            
            // Set worksheet position
            workBook.SetSheetPosition("workSheet2", 0);
            
            workBook.SaveAs("setWorksheetPosition.xlsx");
        }
    }
}