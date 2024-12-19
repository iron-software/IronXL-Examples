using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ClearCells
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.GetWorkSheet("Data");
            
            // Clear cell content
            workSheet["A1"].ClearContents();
            
            workBook.SaveAs("clearSingleCell.xlsx");
        }
    }
}