***Based on <https://ironsoftware.com/examples/excel-formulas-csharp/>***

Utilize IronXL to implement, evaluate, and acquire the computed values through formulas without the need for Office Interop. IronXL currently supports over **150+ formulas** and that number continues to grow with each update. Formulas can be applied using the

`Range.Formula` property. For instance,

```cs
workSheet["A2"].Formula = "=SQRT(A1)";  // Calculates the square root of the value in cell A1
workSheet["B8"].Formula = "=C9/C11";    // Divides the value in cell C9 by the value in cell C11
workSheet["G31"].Formula = "=TAN(G30)"; // Computes the tangent of the angle in cell G30
```

A formula is essentially an expression used to determine the value of a spreadsheet cell. Excel functions, which are predefined formulas, are readily accessible within Excel.

What distinguishes IronXL in the `.NET Excel library` landscape is its robust support for numerous Excel formulas and its ability to instantly compute formula outcomes.

# Implementing Excel Formulas Using C&num;

1. Begin by integrating an Excel library that supports formula functionality.
2. Open the desired Excel file and access the default `Worksheet`.
3. Assign the necessary formulas and values to the selected cells within your spreadsheet.
4. Employ the `EvaluateAll` method to compute all set formulas in the document.
5. Finally, store the changes by saving the `Workbook` object as an Excel file.