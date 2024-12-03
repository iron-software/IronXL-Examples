using IronXL;
using IronXL.Excel;
namespace ironxl.ManageWorksheet
{
    public class Section4
    {
        public void Run()
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