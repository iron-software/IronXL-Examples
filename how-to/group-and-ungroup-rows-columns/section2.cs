using IronXL;

// Load existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Ungroup row 3-5
workSheet.UngroupRows(2, 4);

workBook.SaveAs("ungroupRow.xlsx");