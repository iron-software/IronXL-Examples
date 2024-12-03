using IronXL;
using IronXL.Excel;
namespace ironxl.CopyCells
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");
            
            // Copy cell content
            workSheet["A1"].Copy(workBook.GetWorkSheet("Sheet2"), "B3");
            
            workBook.SaveAs("copyAcrossWorksheet.xlsx");
        }
    }
}