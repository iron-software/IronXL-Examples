using System.Linq;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.SelectRange
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Get range from worksheet
            var range = workSheet["A2:B8"];
        }
    }
}