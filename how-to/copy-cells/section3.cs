using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.CopyCells
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");
            
            // Copy cell content
            workSheet["A1"].Copy(workBook.GetWorkSheet("Sheet2"), "B3");
            
            workBook.SaveAs("copyAcrossWorksheet.xlsx");
        }
    }
}