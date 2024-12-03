using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Set repeating rows for row(2-4)
workSheet.SetRepeatingRows(1, 3);

// Set repeating columns for column(C-D)
workSheet.SetRepeatingColumns(2, 3);

// Set column break after column(H). Hence, the first page will only contain column(A-G)
workSheet.SetColumnBreak(7);

workBook.SaveAs("repeatingRows.xlsx");
