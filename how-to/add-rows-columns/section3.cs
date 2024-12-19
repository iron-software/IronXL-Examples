using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AddRowsColumns
{
    public static class Section3
    {
        public static void Run()
        {
            // Load existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Add a column before column A
            workSheet.InsertColumn(0);
            
            // Insert multiple columns after column B
            workSheet.InsertColumns(2, 2);
            
            workBook.SaveAs("addColumn.xlsx");
        }
    }
}