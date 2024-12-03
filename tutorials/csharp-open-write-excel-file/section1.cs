using IronXL;
using IronXL.Excel;
namespace ironxl.CsharpOpenWriteExcelFile
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("test.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            IronXL.Range range = workSheet["A2:A8"];
            decimal total = 0;
            
            // iterate over range of cells
            foreach (var cell in range)
            {
                Console.WriteLine("Cell {0} has value '{1}'", cell.RowIndex, cell.Value);
                if (cell.IsNumeric)
                {
                    // Get decimal value to avoid floating numbers precision issue
                    total += cell.DecimalValue;
                }
            }
            
            // Check formula evaluation
            if (workSheet["A11"].DecimalValue == total)
            {
                Console.WriteLine("Basic Test Passed");
            }
        }
    }
}