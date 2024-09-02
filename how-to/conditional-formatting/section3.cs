using IronXL;

WorkBook workBook = WorkBook.Load("addConditionalFormatting.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Remove conditional formatting rule
workSheet.ConditionalFormatting.RemoveConditionalFormatting(0);

workBook.SaveAs("removedConditionalFormatting.xlsx");