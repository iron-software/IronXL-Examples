using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.GroupAndUngroupRowsColumns
{
    public static class Section4
    {
        public static void Run()
        {
            // Load existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Ungroup column C-D
            workSheet.UngroupColumn("C", "D");
            
            workBook.SaveAs("ungroupColumn.xlsx");
        }
    }
}