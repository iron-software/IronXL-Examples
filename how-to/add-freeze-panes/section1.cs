using System.Linq;
using IronXL.Excel;
namespace ironxl.AddFreezePanes
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Create freeze pane from column(A-B) and row(1-3)
            workSheet.CreateFreezePane(2, 3);
            
            workBook.SaveAs("createFreezePanes.xlsx");
        }
    }
}