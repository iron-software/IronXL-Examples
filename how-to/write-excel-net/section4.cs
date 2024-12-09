using IronXL;
using IronXL.Excel;
namespace ironxl.WriteExcelNet
{
    public class Section4
    {
        public void Run()
        {
            // Load Excel file
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            
            // Open WorkSheet of sample.xlsx
            WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");
            
            // Access A1 cell and write the value
            workSheet["A1"].Value = "new value";
            
            // Save changes
            workBook.SaveAs("sample.xlsx");
        }
    }
}