using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AutosizeRowsColumns
{
    public static class Section5
    {
        public static void Run()
        {
            // Load existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Apply auto resize to columns individually
            workSheet.AutoSizeColumn(0, true);
            workSheet.AutoSizeColumn(1, true);
            workSheet.AutoSizeColumn(2, true);
            
            workBook.SaveAs("advanceAutoResizeColumn.xlsx");
        }
    }
}