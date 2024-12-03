using IronXL;
using IronXL.Excel;
namespace ironxl.LoadSpreadsheet
{
    public class Section1
    {
        public void Run()
        {
            // Supported for XLSX, XLS, XLSM, XLTX, CSV and TSV
            WorkBook workBook = WorkBook.Load("sample.xlsx");
        }
    }
}