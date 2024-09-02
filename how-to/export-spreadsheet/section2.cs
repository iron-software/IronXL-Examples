using IronXL;

// Create new Excel WorkBook document
WorkBook workBook = WorkBook.Create();

// Create three WorkSheets
WorkSheet workSheet1 = workBook.CreateWorkSheet("sheet1");
WorkSheet workSheet2 = workBook.CreateWorkSheet("sheet2");

// Input information
workSheet1["A1"].StringValue = "A1";
workSheet2["A1"].StringValue = "A1";

// Save as CSV
workBook.SaveAsCsv("sample.csv");

// Save as JSON
workBook.SaveAsJson("sample.json");

// Save as XML
workBook.SaveAsXml("sample.xml");

// Export the excel file as HTML
workBook.ExportToHtml("sample.html");