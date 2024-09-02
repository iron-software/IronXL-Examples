using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Select a column(A)
var column = workSheet.GetColumn(0);

// Sort column(A) in ascending order (A to Z)
column.SortAscending();

// Sort column(A) in descending order (Z to A)
column.SortDescending();

workBook.SaveAs("sortExcelRange.xlsx");