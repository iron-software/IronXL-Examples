using IronXL;

WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Get all named table
var namedTableList = workSheet.GetNamedTableNames();