using IronXL;
using IronXL.Excel;
namespace ironxl.NamedRange
{
    public class Section1
    {
        public void Run()
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