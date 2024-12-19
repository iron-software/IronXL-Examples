using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ExportSpreadsheet
{
    public static class Section1
    {
        public static void Run()
        {
            // Create new Excel WorkBook document
            WorkBook workBook = WorkBook.Create();
            
            // Create a blank WorkSheet
            WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");
            
            // Save the excel file as XLS, XLSX, XLSM, CSV, TSV, JSON, XML, HTML
            workBook.SaveAs("sample.xls");
        }
    }
}