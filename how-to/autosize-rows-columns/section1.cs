using IronXL;

// Load existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Apply auto resize on row 2
workSheet.AutoSizeRow(1);

workBook.SaveAs("autoResize.xlsx");