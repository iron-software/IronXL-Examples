using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.LoadSpreadsheet
{
    public static class Section1
    {
        public static void Run()
        {
            // Supported for XLSX, XLS, XLSM, XLTX, CSV and TSV
            WorkBook workBook = WorkBook.Load("sample.xlsx");
        }
    }
}