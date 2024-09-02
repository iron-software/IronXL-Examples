using IronXL;

// Load workbook
WorkBook workBook = WorkBook.Load("Book1.xlsx");

// Select worksheet
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Edit or Set formula
workSheet["A4"].Formula = "=SUM(A1,A3)";

// Reevaluate the entire workbook
workBook.EvaluateAll();