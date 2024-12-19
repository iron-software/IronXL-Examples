using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.CopyCells
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");
            
            // Copy a single cell(C10)
            workSheet["C10"].Copy(workBook.GetWorkSheet("Sheet1"), "B13");
            
            // Copy a column(A)
            workSheet.GetColumn(0).Copy(workBook.GetWorkSheet("Sheet1"), "H1");
            
            // Copy a row(4)
            workSheet.GetRow(3).Copy(workBook.GetWorkSheet("Sheet1"), "A15");
            
            // Copy a two-dimensional range(D6:F8)
            workSheet["D6:F8"].Copy(workBook.GetWorkSheet("Sheet1"), "H17");
            
            workBook.SaveAs("copyCellRange.xlsx");
        }
    }
}