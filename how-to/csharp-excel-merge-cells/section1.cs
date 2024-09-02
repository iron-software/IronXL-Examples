using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

var range = workSheet["B2:B5"];

// Merge cells B7 to E7
workSheet.Merge("B7:E7");

// Merge selected range
workSheet.Merge(range.RangeAddressAsString);

workBook.SaveAs("mergedCell.xlsx");