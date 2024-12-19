using System.Linq;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.AddFreezePanes
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Create freeze pane from column(A-B) and row(1-3)
            workSheet.CreateFreezePane(2, 3);
            
            workBook.SaveAs("createFreezePanes.xlsx");
        }
    }
}