using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AutosizeRowsColumns
{
    public static class Section7
    {
        public static void Run()
        {
            // Load existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            RangeRow row = workSheet.GetRow(0);
            row.Height = 10; // Set height
            
            RangeColumn col = workSheet.GetColumn(0);
            col.Width = 10; // Set width
            
            workBook.SaveAs("manualHeightAndWidth.xlsx");
        }
    }
}