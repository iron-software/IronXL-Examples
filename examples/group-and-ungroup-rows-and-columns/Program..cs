using IronXL;

// Load existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Apply grouping to row 1-5
workSheet.GroupRows(0, 4);

// Ungroup row 3-5
workSheet.UngroupRows(2, 4);

// Apply grouping to column A-F
workSheet.GroupColumns(0, 5);

// Ungroup column C-D will cut the grouping at B
workSheet.UngroupColumn("C", "D");

workBook.SaveAs("groupAndUngroup.xlsx");
