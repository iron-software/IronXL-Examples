using IronXL;

WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Remove named range
workSheet.RemoveNamedRange("range1");