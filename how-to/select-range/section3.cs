using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Get column from worksheet
var column = workSheet.GetColumn(2);