using IronXL;
using IronXL.Excel;
namespace ironxl.NamedTable
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Get all named table
            var namedTableList = workSheet.GetNamedTableNames();
        }
    }
}