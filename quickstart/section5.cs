using IronXL;
using IronXL.Excel;
namespace ironxl.Quickstart
{
    public class Section5
    {
        public void Run()
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