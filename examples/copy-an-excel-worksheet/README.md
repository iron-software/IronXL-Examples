***Based on <https://ironsoftware.com/examples/copy-an-excel-worksheet/>***

The code snippet presented illustrates the process of using IronXL to clone and replicate `WorkSheets` across different `WorkBooks`, as well as within the same `WorkBook`. This functionality facilitates the transfer of sheets between workbooks or the creation of identical sheets within a single workbook.

To replicate a worksheet within the same workbook or spreadsheet, the `CopySheet` method is employed. This method requires the name of the new worksheet as an argument.

For duplicating a sheet to or from another workbook, utilize the `CopyTo` method. This method requires the `WorkBook` as the initial argument, followed by the name of the new worksheet.

```
// Example of using CopySheet to duplicate a worksheet
var workbook = new WorkBook("example.xlsx");
var originalSheet = workbook.DefaultWorkSheet;
var duplicatedSheet = originalSheet.CopySheet("Copied Sheet");

// Example of using CopyTo to transfer a sheet to another workbook
var targetWorkbook = new WorkBook("target.xlsx");
originalSheet.CopyTo(targetWorkbook, "Transferred Sheet");
```