using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ConditionalFormatting
{
    public static class Section3
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("addConditionalFormatting.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Remove conditional formatting rule
            workSheet.ConditionalFormatting.RemoveConditionalFormatting(0);
            
            workBook.SaveAs("removedConditionalFormatting.xlsx");
        }
    }
}