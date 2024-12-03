using System.Linq;
using IronXL.Excel;
namespace ironxl.SelectRange
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Get range from worksheet
            var range = workSheet["A2:B8"];
        }
    }
}