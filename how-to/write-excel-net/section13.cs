using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Write new above old in complete WorkSheet
workSheet.Replace("old", "new");

// Write new above old just in row no 6 of WorkSheet
workSheet.Rows[5].Replace("old", "new");

// Write new above old just in column no 5 of WorkSheet
workSheet.Columns[4].Replace("old", "new");

// Write new above old just from A5 to H5 of WorkSheet
workSheet["A5:H5"].Replace("old", "new");

workBook.SaveAs("sample.xlsx");