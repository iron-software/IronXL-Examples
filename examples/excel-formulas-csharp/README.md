> Full guide: [Excel formulas C#](https://ironsoftware.com/csharp/excel/examples/excel-formulas-csharp/?utm_source=github)

Utilize IronXL to implement, evaluate, and acquire the computed values through formulas without the need for Office Interop. IronXL currently supports over **150+ formulas** and that number continues to grow with each update. Formulas can be applied using the

`Range.Formula` property. For instance,

```cs
workSheet["A2"].Formula = "=SQRT(A1)";  // Calculates the square root of the value in cell A1
workSheet["B8"].Formula = "=C9/C11";    // Divides the value in cell C9 by the value in cell C11
workSheet["G31"].Formula = "=TAN(G30)"; // Computes the tangent of the angle in cell G30
```

A formula is essentially an expression used to determine the value of a spreadsheet cell. Excel functions, which are predefined formulas, are readily accessible within Excel.

IronXL supports a wide range of Excel formulas and computes their results immediately.

# Implementing Excel Formulas Using C&num;

1. Begin by integrating an Excel library that supports formula functionality.
2. Open the desired Excel file and access the default `Worksheet`.
3. Assign the necessary formulas and values to the selected cells within your spreadsheet.
4. Employ the `EvaluateAll` method to compute all set formulas in the document.
5. Finally, store the changes by saving the `Workbook` object as an Excel file.