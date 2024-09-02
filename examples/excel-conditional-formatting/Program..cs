using IronXL;
using IronXL.Formatting.Enums;
using IronXL.Styles;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Create conditional formatting rule
var rule = workSheet.ConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.LessThan, "8");
// Set style options
rule.FontFormatting.IsBold = true;
rule.FontFormatting.FontColor = "#123456";
rule.BorderFormatting.RightBorderColor = "#ffffff";
rule.BorderFormatting.RightBorderType = BorderType.Thick;
rule.PatternFormatting.BackgroundColor = "#54bdd9";
rule.PatternFormatting.FillPattern = FillPattern.Diamonds;
// Apply formatting on specified region
workSheet.ConditionalFormatting.AddConditionalFormatting("A3:A8", rule);

// Create conditional formatting rule
var rule1 = workSheet.ConditionalFormatting.CreateConditionalFormattingRule(ComparisonOperator.Between, "7", "10");
// Set style options
rule1.FontFormatting.IsItalic = true;
rule1.FontFormatting.UnderlineType = FontUnderlineType.Single;
// Apply formatting on specified region
workSheet.ConditionalFormatting.AddConditionalFormatting("A3:A9", rule1);

workBook.SaveAs("applyConditionalFormatting.xlsx");
