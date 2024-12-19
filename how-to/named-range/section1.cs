using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.NamedRange
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Create();
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Select range
            var selectedRange = workSheet["A1:A5"];
            
            // Add named range
            workSheet.AddNamedRange("range1", selectedRange);
            
            workBook.SaveAs("addNamedRange.xlsx");
        }
    }
}