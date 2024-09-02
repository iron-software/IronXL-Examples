using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Select a range
var range = workSheet["A1:D20"];

// Select a column(B)
var column = workSheet.GetColumn(1);

// Sort the range in ascending order (A to Z)
range.SortAscending();

// Sort the range by column(C) in ascending order
range.SortByColumn("C", SortOrder.Ascending);

// Sort the column(B) in descending order (Z to A)
column.SortDescending();

workBook.SaveAs("sortExcelRange.xlsx");
