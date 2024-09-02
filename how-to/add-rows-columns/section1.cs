using IronXL;

// Load existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Add a row before row 2
workSheet.InsertRow(1);

// Insert multiple rows after row 3
workSheet.InsertRows(3, 3);

workBook.SaveAs("addRow.xlsx");