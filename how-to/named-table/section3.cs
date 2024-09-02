using IronXL;

WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Get named table
var namedRangeAddress = workSheet.GetNamedTable("table1");