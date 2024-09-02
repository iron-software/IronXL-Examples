using IronXL;

// Create new Excel WorkBook document
WorkBook workBook = WorkBook.Create();

// Create a blank WorkSheet
WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

// Save the excel file as XLS, XLSX, XLSM, CSV, TSV, JSON, XML, HTML
workBook.SaveAs("sample.xls");