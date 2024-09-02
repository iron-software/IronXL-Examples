using IronXL;

// Load existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Apply auto resize on column A
workSheet.AutoSizeColumn(0);

workBook.SaveAs("autoResizeColumn.xlsx");