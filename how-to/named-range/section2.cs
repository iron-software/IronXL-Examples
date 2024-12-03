using IronXL;
using IronXL.Excel;
namespace ironxl.NamedRange
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Get all named range
            var namedRangeList = workSheet.GetNamedRanges();
        }
    }
}