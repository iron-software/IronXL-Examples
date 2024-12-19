using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.Overview.Quickstart
{
    public static class Section1
    {
        public static void Run()
        {
            // Supported spreadsheet formats for reading include: XLSX, XLS, CSV and TSV
            WorkBook workBook = WorkBook.Load("data.xlsx");
            WorkSheet workSheet = workBook.WorkSheets.First();
            
            // Select cells easily in Excel notation and return the calculated value, date, text or formula
            int cellValue = workSheet["A2"].IntValue;
            
            // Read from Ranges of cells elegantly.
            foreach (var cell in workSheet["A2:B10"])
            {
                Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
            }
        }
    }
}