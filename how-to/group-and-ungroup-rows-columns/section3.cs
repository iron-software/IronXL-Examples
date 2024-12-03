using IronXL;
using IronXL.Excel;
namespace ironxl.GroupAndUngroupRowsColumns
{
    public class Section3
    {
        public void Run()
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