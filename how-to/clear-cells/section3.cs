using IronXL;
using IronXL.Excel;
namespace ironxl.ClearCells
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            
            // Delete all worksheets
            workBook.WorkSheets.Clear();
            
            workBook.SaveAs("useClear.xlsx");
        }
    }
}