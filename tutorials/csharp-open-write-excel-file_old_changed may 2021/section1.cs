using System;
using IronXL.Excel;
namespace IronXL.Examples.Tutorial.CsharpOpenWriteExcelFile_Old_Changed May 2021
{
    public static class Section1
    {
        public static void Run()
        {
            WorkBook workbook = WorkBook.Load("test.xlsx");
            WorkSheet sheet = workbook.DefaultWorkSheet;
            
            Range range = sheet ["A2:A8"];
            
            decimal total = 0;
            
            //iterate over range of cells
            foreach (var cell in range)
            {
                Console.WriteLine("Cell {0} has value '{1}'", cell.RowIndex, cell.Value);
            
                if (cell.IsNumeric)
                {
                    //Get decimal value to avoid floating numbers precision issue
                    total += cell.DecimalValue;
                }
            }
            
            //check formula evaluation
            if (sheet ["A11"].DecimalValue == total)
            {
                Console.WriteLine("Basic Test Passed");
            }
        }
    }
}