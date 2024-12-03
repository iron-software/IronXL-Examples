using IronXL;
using IronXL.Excel;
namespace ironxl.CreateSpreadsheet
{
    public class Section2
    {
        public void Run()
        {
            // Create XLSX spreadsheet
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
        }
    }
}