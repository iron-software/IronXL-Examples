using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.GroupAndUngroupRowsColumns
{
    public static class Section3
    {
        public static void Run()
        {
            // Load existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Apply grouping to column A-F
            workSheet.GroupColumns(0, 5);
            
            workBook.SaveAs("groupColumn.xlsx");
        }
    }
}