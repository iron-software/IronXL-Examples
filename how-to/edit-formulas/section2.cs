using System.Linq;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.EditFormulas
{
    public static class Section2
    {
        public static void Run()
        {
            // Load workbook
            WorkBook workBook = WorkBook.Load("Book1.xlsx");
            
            // Select worksheet
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Retrieve the result value
            string value = workSheet["A4"].First().FormattedCellValue;
            
            // Print the result to console
            Console.WriteLine(value);
        }
    }
}