using IronXL.Formatting.Enums;
using IronXL.Excel;
namespace ironxl.ConditionalFormatting
{
    public class Section1
    {
        public void Run()
        {
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Create conditional formatting rule
            var rule = workSheet.ConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.LessThan, "8");
            
            // Set style options
            rule.PatternFormatting.BackgroundColor = "#54BDD9";
            
            // Add conditional formatting rule
            workSheet.ConditionalFormatting.AddConditionalFormatting("A1:A10", rule);
            
            workBook.SaveAs("addConditionalFormatting.xlsx");
        }
    }
}