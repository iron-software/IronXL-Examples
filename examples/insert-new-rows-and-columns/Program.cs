using IronXL;

// Load existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Add a row before row 1
workSheet.InsertRow(0);

// Insert multiple rows before row 4
workSheet.InsertRows(3, 3);

// Add a column before column A
workSheet.InsertColumn(0);

// Insert multiple columns before column F
workSheet.InsertColumns(5, 2);

workBook.SaveAs("addRowAndColumn.xlsx");
