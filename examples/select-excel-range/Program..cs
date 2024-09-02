using IronXL;
using System;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Get range from worksheet
var range = workSheet["A2:A8"];

// Get column from worksheet
var columnA = workSheet.GetColumn(0);

// Get row from worksheet
var row1 = workSheet.GetRow(0);

// Iterate over the range
foreach (var cell in range)
{
    Console.WriteLine($"{cell.Value}");
}

// Select and print every row
var rows = workSheet.Rows;

foreach (var eachRow in rows)
{
    foreach (var cell in eachRow)
    {
        Console.Write($"  {cell.Value}  |");
    }
    Console.WriteLine($"");
}
