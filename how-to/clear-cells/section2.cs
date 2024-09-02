using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Data");

// Clear a single cell(A1)
workSheet["A1"].ClearContents();

// Clear a column(B)
workSheet.GetColumn("B").ClearContents();

// Clear a row(4)
workSheet.GetRow(3).ClearContents();

// Clear a two-dimensional range(D6:F9)
workSheet["D6:F9"].ClearContents();

workBook.SaveAs("clearCellRange.xlsx");