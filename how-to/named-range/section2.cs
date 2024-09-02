using IronXL;

WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Get all named range
var namedRangeList = workSheet.GetNamedRanges();