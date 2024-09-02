using IronXL;

WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Get named range address
string namedRangeAddress = workSheet.FindNamedRange("range1");

// Select range
var range = workSheet[$"{namedRangeAddress}"];