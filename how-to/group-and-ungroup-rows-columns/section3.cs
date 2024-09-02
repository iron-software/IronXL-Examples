using IronXL;

// Load existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Apply grouping to column A-F
workSheet.GroupColumns(0, 5);

workBook.SaveAs("groupColumn.xlsx");