using System.Linq;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.SelectRange
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Get column from worksheet
            var column = workSheet.GetColumn(2);
        }
    }
}