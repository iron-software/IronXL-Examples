using IronXL;
using IronXL.Excel;
namespace ironxl.AutosizeRowsColumns
{
    public class Section1
    {
        public void Run()
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