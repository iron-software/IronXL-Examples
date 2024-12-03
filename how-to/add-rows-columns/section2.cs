using IronXL;
using IronXL.Excel;
namespace ironxl.AddRowsColumns
{
    public class Section2
    {
        public void Run()
        {
            // Load existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Remove row 5
            workSheet.GetRow(4).RemoveRow();
            
            workBook.SaveAs("removeRow.xlsx");
        }
    }
}