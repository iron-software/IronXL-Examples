using IronXL;
using IronXL.Excel;
namespace IronXL.Examples.HowTo.ConditionalFormatting
{
    public static class Section2
    {
        public static void Run()
        {
            WorkBook workBook = WorkBook.Load("addConditionalFormatting.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Create conditional formatting rule
            var ruleCollection = workSheet.ConditionalFormatting.GetConditionalFormattingAt(0);
            var rule = ruleCollection.GetRule(0);
            
            // Edit styling
            rule.PatternFormatting.BackgroundColor = "#B6CFB6";
            
            workBook.SaveAs("editedConditionalFormatting.xlsx");
        }
    }
}