using IronXL;
using IronXL.Excel;
namespace ironxl.AutosizeRowsColumns
{
    public class Section3
    {
        public void Run()
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