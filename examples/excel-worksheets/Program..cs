using IronXL;

// Create new Excel spreadsheet
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);

// Create worksheets (workSheet1, workSheet2, workSheet3)
WorkSheet workSheet1 = workBook.CreateWorkSheet("workSheet1");
WorkSheet workSheet2 = workBook.CreateWorkSheet("workSheet2");
WorkSheet workSheet3 = workBook.CreateWorkSheet("workSheet3");

// Set worksheet position (workSheet2, workSheet1, workSheet3)
workBook.SetSheetPosition("workSheet2", 0);

// Set active for workSheet3
workBook.SetActiveTab(2);

// Remove workSheet1
workBook.RemoveWorkSheet(1);

workBook.SaveAs("manageWorkSheet.xlsx");
