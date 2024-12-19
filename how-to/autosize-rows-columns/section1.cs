using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AutosizeRowsColumns
{
    public static class Section1
    {
        public static void Run()
        {
            // Load existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Apply auto resize on row 2
            workSheet.AutoSizeRow(1);
            
            workBook.SaveAs("autoResize.xlsx");
        }
    }
}