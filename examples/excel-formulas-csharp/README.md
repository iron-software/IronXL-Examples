***Based on <https://ironsoftware.com/examples/excel-formulas-csharp/>***

Utilize IronXL to input formulas, assess, and obtain computed output values without needing Office Interop. IronXL currently supports more than **150+ formulas**, with more being added as the software evolves. You can specify a formula using the `Range.Formula` property. Here's how you can do it:

```cs
workSheet["A2"].Formula = "=SQRT(A1)";
workSheet["B8"].Formula = "=C9/C11";
workSheet["G31"].Formula = "=TAN(G30)";
```

A formula is simply an expression that determines the value of a cell. Functions, which are essentially built-in formulas, are readily accessible in Excel.

IronXL distinguishes itself among .NET Excel libraries due to its broad support for various Excel formulas and its ability to instantly calculate the values of these formulas.

# Implementing Excel Formulas in C&num;

1. Install a library capable of handling Excel formulas.
2. Open the Excel document and access the default `Worksheet`.
3. Assign formulas and values to the cells in the spreadsheet as needed.
4. Utilize the `EvaluateAll` method to compute all set formulas.
5. Save the updated `Workbook` to an Excel file.