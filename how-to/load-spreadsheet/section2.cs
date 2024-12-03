using IronXL;
using IronXL.Excel;
namespace ironxl.LoadSpreadsheet
{
    public class Section2
    {
        public void Run()
        {
            // Load CSV file
            WorkBook workBook = WorkBook.LoadCSV("sample.csv");
        }
    }
}