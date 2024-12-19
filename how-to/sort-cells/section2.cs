using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.SortCells
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Select a range
            var range = workSheet["A1:D10"];
            
            // Sort the range by column(B) in ascending order
            range.SortByColumn("B", SortOrder.Ascending);
            
            workBook.SaveAs("sortRange.xlsx");
        }
    }
}