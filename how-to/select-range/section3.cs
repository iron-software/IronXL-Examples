using System.Linq;
using IronXL.Excel;
namespace ironxl.SelectRange
{
    public class Section3
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Get column from worksheet
            var column = workSheet.GetColumn(2);
        }
    }
}