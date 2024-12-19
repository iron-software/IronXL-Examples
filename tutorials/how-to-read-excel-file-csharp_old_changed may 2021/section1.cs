using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp_Old_Changed May 2021
{
    public static class Section1
    {
        public static void Run()
        {
            /**
            Read XLS or XLSX File
            anchor-read-an-xls-or-xlsx-file
            **/
            using IronXL;
            using System.Linq;
                
            //Supported spreadsheet formats for reading include: XLSX, XLS, CSV and TSV
            WorkBook workbook = WorkBook.Load("test.xlsx");
            WorkSheet sheet = workbook.WorkSheets.First();
            //Select cells easily in Excel notation and return the calculated value
            int cellValue = sheet ["A2"].IntValue;
            // Read from Ranges of cells elegantly.
            foreach (var cell in sheet ["A2:A10"])
            {
                Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
            }
            
            
            ///Advanced Operations
            
            //Calculate aggregate values such as Min, Max and Sum
            decimal sum = sheet ["A2:A10"].Sum();
            //Linq compatible
            decimal max = sheet ["A2:A10"].Max(c => c.DecimalValue);
        }
    }
}