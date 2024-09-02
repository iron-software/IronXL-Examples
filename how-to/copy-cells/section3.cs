using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Copy cell content
workSheet["A1"].Copy(workBook.GetWorkSheet("Sheet2"), "B3");

workBook.SaveAs("copyAcrossWorksheet.xlsx");