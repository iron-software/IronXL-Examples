using IronXL;

// Create new Excel WorkBook document
WorkBook workBook = WorkBook.Create();

// Create a blank WorkSheet
WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

// Add data and styles to the new worksheet
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");

// Save the excel file as XLS, XLSX, XLSM, CSV, TSV, JSON, XML, HTML
workBook.SaveAs("sample.xls");
workBook.SaveAs("sample.xlsx");
workBook.SaveAs("sample.tsv");

// Save the excel file as CSV
workBook.SaveAsCsv("sample.csv");

// Save the excel file as JSON
workBook.SaveAsJson("sample.json");

// Save the excel file as XML
workBook.SaveAsXml("sample.xml");

// Export the excel file as HTML
workBook.ExportToHtml("sample.html");
