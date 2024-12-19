using IronXL.Excel;
namespace IronXL.Examples.Tutorial.HowToReadExcelFileCsharp_Old_Changed May 2021
{
    public static class Section11
    {
        public static void Run()
        {
            /**
            Add Spreadsheet Formulae
            anchor-add-formulae-to-a-spreadsheet
            **/
            //Iterate through all rows with a value
            for (var y = 2; y < i; y++)
            {
                //Get the C cell
                var cell = sheet [$"C{y}"].First();
            
                //Set the formula for the Percentage of Total column
                cell.Formula = $"=B{y}/B{i}";
            }
        }
    }
}