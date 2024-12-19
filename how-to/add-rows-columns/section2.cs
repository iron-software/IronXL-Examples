using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AddRowsColumns
{
    public static class Section2
    {
        public static void Run()
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