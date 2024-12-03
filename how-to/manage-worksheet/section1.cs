using IronXL;
using IronXL.Excel;
namespace ironxl.ManageWorksheet
{
    public class Section1
    {
        public void Run()
        {
            // Create new Excel spreadsheet
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
            
            // Create worksheets
            WorkSheet workSheet1 = workBook.CreateWorkSheet("workSheet1");
            WorkSheet workSheet2 = workBook.CreateWorkSheet("workSheet2");
            WorkSheet workSheet3 = workBook.CreateWorkSheet("workSheet3");
            WorkSheet workSheet4 = workBook.CreateWorkSheet("workSheet4");
            
            
            workBook.SaveAs("createNewWorkSheets.xlsx");
        }
    }
}