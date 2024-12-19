using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.NamedRange
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Get all named range
            var namedRangeList = workSheet.GetNamedRanges();
        }
    }
}