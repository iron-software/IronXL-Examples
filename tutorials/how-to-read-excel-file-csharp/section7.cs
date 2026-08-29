using IronXL;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp
{
    public static class Section7
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;

            IronXL.Cell cell = workSheet["B1"].First();
            string value = cell.StringValue;   // Read the value of the cell as a string
            Console.WriteLine(value);
            
            cell.Value = "10.3289";           // Write a new value to the cell
            Console.WriteLine(cell.StringValue);
        }
    }
}