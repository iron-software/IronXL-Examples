using IronXL;

WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");

// Remove workSheet1
workBook.RemoveWorkSheet(1);

// Remove workSheet2
workBook.RemoveWorkSheet("workSheet2");

workBook.SaveAs("removeWorksheet.xlsx");