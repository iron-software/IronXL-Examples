using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Data");

// Clear cell content
workSheet["A1"].ClearContents();

workBook.SaveAs("clearSingleCell.xlsx");