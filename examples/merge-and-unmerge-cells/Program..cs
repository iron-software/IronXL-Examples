using IronXL;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.DefaultWorkSheet;

var range = workSheet["B1:B3"];

// Merge cells D1 to D3
workSheet.Merge("D1:D3");

// Merge selected range
workSheet.Merge(range.RangeAddressAsString);

workBook.SaveAs("mergedCell.xlsx");

// Unmerge the merged region of D1 to D3
workSheet.Unmerge("D1:D3");
