using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AutosizeRowsColumns
{
    public static class Section3
    {
        public static void Run()
        {
            // Load existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Apply auto resize to rows individually
            workSheet.AutoSizeRow(0, true);
            workSheet.AutoSizeRow(1, true);
            workSheet.AutoSizeRow(2, true);
            
            workBook.SaveAs("advanceAutoResizeRow.xlsx");
        }
    }
}