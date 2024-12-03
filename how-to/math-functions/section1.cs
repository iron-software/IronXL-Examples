using System.Linq;
using IronXL.Excel;
namespace ironxl.MathFunctions
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Get range from worksheet
            var range = workSheet["A1:A8"];
            
            // Calculate the sum of numeric cells within the range
            decimal sum = range.Sum();
            
            // Calculate the average value of numeric cells within the range
            decimal avg = range.Avg();
            
            // Identify the maximum value among numeric cells within the range
            decimal max = range.Max();
            
            // Identify the minimum value among numeric cells within the range
            decimal min = range.Min();
        }
    }
}