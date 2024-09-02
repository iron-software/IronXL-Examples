using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");

// Delete all worksheets
workBook.WorkSheets.Clear();

workBook.SaveAs("useClear.xlsx");