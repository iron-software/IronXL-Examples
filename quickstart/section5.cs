using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.Overview.Quickstart
{
    public static class Section5
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("test.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // This is how we get range from Excel worksheet
            Range range = workSheet["A2:A8"];
            
            // Sort the range in the sheet
            range.SortAscending();
            workBook.Save();
        }
    }
}