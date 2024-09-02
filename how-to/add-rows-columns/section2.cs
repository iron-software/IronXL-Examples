using IronXL;

// Load existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Remove row 5
workSheet.GetRow(4).RemoveRow();

workBook.SaveAs("removeRow.xlsx");