***Based on <https://ironsoftware.com/examples/copy-an-excel-worksheet/>***

The preceding code sample demonstrates leveraging IronXL to replicate and transfer `WorkSheets` across different `WorkBooks`, and even within the same `WorkBook`. This enables the copying and pasting of sheets between workbooks or duplicating them in the same workbook.

For duplicating a worksheet within the same workbook or spreadsheet, you would utilize the `CopySheet` method. This function necessitates specifying the name of the new worksheet as a parameter.

To replicate a sheet to another workbook or from another workbook, the `CopyTo` method is utilized. This requires specifying the `WorkBook` as the first parameter followed by the name of the new worksheet.