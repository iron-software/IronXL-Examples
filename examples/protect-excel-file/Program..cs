using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Open protected spreadsheet file
WorkBook protectedWorkBook = WorkBook.Load("sample.xlsx", "IronSoftware");

// Spreadsheet protection
// Set protection for spreadsheet file
workBook.Encrypt("IronSoftware");

// Remove protection for spreadsheet file. Original password is required.
workBook.Password = null;

workBook.Save();

// Worksheet protection
// Set protection for individual worksheet
workSheet.ProtectSheet("IronXL");

// Remove protection for particular worksheet. It works without password!
workSheet.UnprotectSheet();

workBook.Save();
