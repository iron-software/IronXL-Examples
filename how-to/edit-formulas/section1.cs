using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.EditFormulas
{
    public static class Section1
    {
        public static void Run()
        {
            // Load workbook
            WorkBook workBook = WorkBook.Load("Book1.xlsx");
            
            // Select worksheet
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Edit or Set formula
            workSheet["A4"].Formula = "=SUM(A1,A3)";
            
            // Reevaluate the entire workbook
            workBook.EvaluateAll();
        }
    }
}