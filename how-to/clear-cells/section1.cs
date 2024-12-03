using IronXL;
using IronXL.Excel;
namespace ironxl.ClearCells
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.GetWorkSheet("Data");
            
            // Clear cell content
            workSheet["A1"].ClearContents();
            
            workBook.SaveAs("clearSingleCell.xlsx");
        }
    }
}