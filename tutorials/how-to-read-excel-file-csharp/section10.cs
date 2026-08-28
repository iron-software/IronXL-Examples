using IronXL;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp
{
    public static class Section10
    {
        public static void Run()
        {
            // The docs page opens the workbook before this snippet; declared here so
            // the section stands on its own.
            IronXL.WorkBook workBook = IronXL.WorkBook.Load("sample.xlsx");
            IronXL.WorkSheet workSheet = workBook.DefaultWorkSheet;
            int i = 0;

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