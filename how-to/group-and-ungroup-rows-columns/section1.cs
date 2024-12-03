using IronXL;
using IronXL.Excel;
namespace ironxl.GroupAndUngroupRowsColumns
{
    public class Section1
    {
        public void Run()
        {
            // Load existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Ungroup row 1-9
            workSheet.GroupRows(0, 7);
            
            workBook.SaveAs("groupRow.xlsx");
        }
    }
}