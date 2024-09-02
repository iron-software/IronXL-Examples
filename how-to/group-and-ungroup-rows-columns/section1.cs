using IronXL;

// Load existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Ungroup row 1-9
workSheet.GroupRows(0, 7);

workBook.SaveAs("groupRow.xlsx");