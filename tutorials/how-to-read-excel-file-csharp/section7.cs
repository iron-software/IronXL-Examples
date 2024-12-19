using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp
{
    public static class Section7
    {
        public static void Run()
        {
            IronXL.Cell cell = workSheet["B1"].First();
            string value = cell.StringValue;   // Read the value of the cell as a string
            Console.WriteLine(value);
            
            cell.Value = "10.3289";           // Write a new value to the cell
            Console.WriteLine(cell.StringValue);
        }
    }
}