using IronXL;
using IronXL.Excel;
namespace ironxl.GroupAndUngroupRowsColumns
{
    public class Section4
    {
        public void Run()
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