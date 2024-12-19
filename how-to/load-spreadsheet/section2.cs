using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.LoadSpreadsheet
{
    public static class Section2
    {
        public static void Run()
        {
            // Load CSV file
            WorkBook workBook = WorkBook.LoadCSV("sample.csv");
        }
    }
}