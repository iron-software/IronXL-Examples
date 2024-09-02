using IronXL;

// Load Excel file
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Open WorkSheet of sample.xlsx
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Specify range row wise and write new value
workSheet["B2:B9"].Value = "new value";

// Specify range column wise and write new value
workSheet["C3:C7"].Value = "new value";

// Save changes
workBook.SaveAs("sample.xlsx");