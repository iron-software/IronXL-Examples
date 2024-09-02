using IronXL;
using System;
using System.Linq;

// Supported for XLSX, XLS, XLSM, XLTX, CSV and TSV
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Select worksheet at index 0
WorkSheet workSheet = workBook.WorkSheets[0];

// Get any existing worksheet
WorkSheet firstSheet = workBook.DefaultWorkSheet;

// Select a cell and return the converted value
int cellValue = workSheet["A2"].IntValue;

// Read from ranges of cells elegantly.
foreach (var cell in workSheet["A2:A10"])
{
    Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
}

// Calculate aggregate values such as Min, Max and Sum
decimal sum = workSheet["A2:A10"].Sum();

// Linq compatible
decimal max = workSheet["A2:A10"].Max(c => c.DecimalValue);
