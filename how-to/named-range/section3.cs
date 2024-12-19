using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.NamedRange
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Get named range address
            string namedRangeAddress = workSheet.FindNamedRange("range1");
            
            // Select range
            var range = workSheet[$"{namedRangeAddress}"];
        }
    }
}