using IronXL;
using IronXL.Excel;
namespace ironxl.NamedRange
{
    public class Section4
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Remove named range
            workSheet.RemoveNamedRange("range1");
        }
    }
}