using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.NamedTable
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Get all named table
            var namedTableList = workSheet.GetNamedTableNames();
        }
    }
}