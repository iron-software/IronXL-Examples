using IronXL;
using System;
using System.Linq;

// Supported spreadsheet formats for reading include: XLSX, XLS, CSV and TSV
WorkBook workBook = WorkBook.Load("test.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Select cells easily in Excel notation and return the calculated value
int cellValue = workSheet["A2"].IntValue;

// Read from Ranges of cells elegantly.
foreach (var cell in workSheet["A2:A10"])
{
    Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
}

// Advanced Operations
// Calculate aggregate values such as Min, Max and Sum
decimal sum = workSheet["A2:A10"].Sum();

// Linq compatible
decimal max = workSheet["A2:A10"].Max(c => c.DecimalValue);