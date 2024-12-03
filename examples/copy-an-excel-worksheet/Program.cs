using IronXL;

WorkBook firstBook = WorkBook.Load("sample.xlsx");
WorkBook secondBook = WorkBook.Create();

// Select first worksheet in the workbook
WorkSheet workSheet = firstBook.DefaultWorkSheet;

// Duplicate the worksheet to the same workbook
workSheet.CopySheet("Copied Sheet");

// Duplicate the worksheet to another workbook with the specified name
workSheet.CopyTo(secondBook, "Copied Sheet");

firstBook.Save();
secondBook.SaveAs("copyExcelWorksheet.xlsx");
