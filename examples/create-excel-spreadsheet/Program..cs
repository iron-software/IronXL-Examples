using IronXL;

// Create new Excel WorkBook document
WorkBook workBook = WorkBook.Create();

// Convert XLSX to XLS
WorkBook xlsWorkBook = WorkBook.Create(ExcelFileFormat.XLS);

// Create a blank WorkSheet
WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

// Add data and styles to the new worksheet
workSheet["A1"].Value = "Hello World";
workSheet["A1"].Style.WrapText = true;
workSheet["A2"].BoolValue = true;
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;

// Save the excel file as XLS, XLSX, CSV, TSV, JSON, XML, HTML and streams
workBook.SaveAs("sample.xlsx");
