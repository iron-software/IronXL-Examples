using IronXL;

// Load existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Apply auto resize on row 2
workSheet.AutoSizeRow(1);

// Apply auto resize on column A
workSheet.AutoSizeColumn(0);

// Apply auto resize on column D
workSheet.AutoSizeColumn(3, true);

workBook.SaveAs("autoResize.xlsx");
