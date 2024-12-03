using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Set Formulas
workSheet["A1"].Formula = "Sum(B8:C12)";
workSheet["B8"].Formula = "=C9/C11";
workSheet["G30"].Formula = "Max(C3:C7)";

// Force recalculate all formula values in all sheets.
workBook.EvaluateAll();

// Get the formula's calculated value.  e.g. "52"
var formulaValue = workSheet["G30"].First().FormattedCellValue;

// Get the formula as a string. e.g. "Max(C3:C7)"
string formulaString = workSheet["G30"].Formula;

// Save changes with updated formulas and calculated values.
workBook.Save();
