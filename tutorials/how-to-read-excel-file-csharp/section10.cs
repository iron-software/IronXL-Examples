using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp
{
    public static class Section10
    {
        public static void Run()
        {
            // Iterate through all rows with a value
            for (var y = 2 ; y < i ; y++)
            {
                // Get the C cell
                Cell cell = workSheet[$"C{y}"].First();
            
                // Set the formula for the Percentage of Total column
                cell.Formula = $"=B{y}/B{i}";
            }
        }
    }
}