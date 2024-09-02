WorkBook workBook = WorkBook.Load("test.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;
IronXL.Cell cell = workSheet["B1"].First();