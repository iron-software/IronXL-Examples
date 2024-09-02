using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Get range from worksheet
var range = workSheet["A2:B8"];