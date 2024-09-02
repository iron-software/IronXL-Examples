using IronXL;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Select range
var selectedRange = workSheet["A1:A5"];

// Add named range
workSheet.AddNamedRange("range1", selectedRange);

workBook.SaveAs("addNamedRange.xlsx");