using IronXL;
using System;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Get a range from an Excel worksheet
var range = workSheet["A2:A8"];

// Combine two ranges
var combinedRange = range + workSheet["A9:A10"];

// Iterate over combined range
foreach (var cell in combinedRange)
{
    Console.WriteLine(cell.Value);
}
