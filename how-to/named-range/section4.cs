using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.NamedRange
{
    public static class Section4
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Remove named range
            workSheet.RemoveNamedRange("range1");
        }
    }
}