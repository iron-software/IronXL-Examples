using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.NamedTable
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Get named table
            var namedRangeAddress = workSheet.GetNamedTable("table1");
        }
    }
}