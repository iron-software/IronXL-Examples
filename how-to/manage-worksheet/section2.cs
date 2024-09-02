using IronXL;

WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");

// Set worksheet position
workBook.SetSheetPosition("workSheet2", 0);

workBook.SaveAs("setWorksheetPosition.xlsx");