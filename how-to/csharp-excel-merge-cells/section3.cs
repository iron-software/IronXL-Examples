using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.CsharpExcelMergeCells
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("mergedCell.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Unmerge the merged region of B7 to E7
            workSheet.Unmerge("B7:E7");
            
            workBook.SaveAs("unmergedCell.xlsx");
        }
    }
}