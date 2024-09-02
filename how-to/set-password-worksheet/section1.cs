using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Set protection for selected worksheet
workSheet.ProtectSheet("IronXL");

workBook.Save();