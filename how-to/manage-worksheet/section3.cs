using IronXL;

WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");

// Set active for workSheet3
workBook.SetActiveTab(2);

workBook.SaveAs("setActiveTab.xlsx");