using IronXL;
using IronXL.Excel;
namespace ironxl.NamedTable
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Get named table
            var namedRangeAddress = workSheet.GetNamedTable("table1");
        }
    }
}