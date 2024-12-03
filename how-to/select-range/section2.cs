using System.Linq;
using IronXL.Excel;
namespace ironxl.SelectRange
{
    public class Section2
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xls");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Get row from worksheet
            var row = workSheet.GetRow(3);
        }
    }
}