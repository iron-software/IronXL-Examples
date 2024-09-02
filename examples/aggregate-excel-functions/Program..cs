using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Get range from worksheet
var range = workSheet["A1:A8"];

// Apply sum of all numeric cells within the range
decimal sum = range.Sum();

// Apply average value of all numeric cells within the range
decimal avg = range.Avg();

// Identify maximum value of all numeric cells within the range
decimal max = range.Max();

// Identify minimum value of all numeric cells within the range
decimal min = range.Min();
