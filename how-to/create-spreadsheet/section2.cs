using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.CreateSpreadsheet
{
    public static class Section2
    {
        public static void Run()
        {
            // Create XLSX spreadsheet
            WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
        }
    }
}