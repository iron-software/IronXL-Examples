using System.Linq;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.SelectRange
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Get row from worksheet
            var row = workSheet.GetRow(3);
        }
    }
}