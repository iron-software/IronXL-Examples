# Managing Worksheets with IronXL

***Based on <https://ironsoftware.com/how-to/manage-worksheet/>***


The IronXL library provides a straightforward approach to managing worksheets within your C# applications. This powerful tool enables you to create, delete, reposition, and set the active worksheet in an Excel file, completely eliminating the need for Office Interop.

***

### Getting Started with IronXL

***

## Example of Worksheet Management

IronXL facilitates efficient worksheet management, allowing you to create, move, and delete worksheets seamlessly using concise code syntax.

Indices used here are zero-based.

## Creating a Worksheet

To add a new worksheet, use the `CreateWorksheet` method, which only requires the desired name of the worksheet. Once created, additional operations such as [merging cells](https://ironsoftware.com/csharp/excel/how-to/csharp-excel-merge-cells/) can be immediately applied.

```cs
using IronXL;

// Initialize a new Excel workbook
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);

// Adding multiple worksheets
WorkSheet workSheet1 = workBook.CreateWorkSheet("workSheet1");
WorkSheet workSheet2 = workBook.CreateWorkSheet("workSheet2");
WorkSheet workSheet3 = workBook.CreateWorkSheet("workSheet3");
WorkSheet workSheet4 = workBook.CreateWorkSheet("workSheet4");


workBook.SaveAs("createNewWorkSheets.xlsx");
```

<center>
    ![Create Worksheets](https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-create-worksheet.png)
</center>

<hr>

## Reordering Worksheets
The method `SetSheetPosition` relocates a worksheet to a new position within the workbook. It requires the worksheet’s name and new index position.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");

// Reorder the second worksheet to the first position
workBook.SetSheetPosition("workSheet2", 0);

workBook.SaveAs("setWorksheetPosition.xlsx");
```

<center>
    ![Change Worksheet Position](https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-set-sheet-position.png)
</center>

<hr>

## Activating a Worksheet
Use the `SetActiveTab` method to specify the default active worksheet when the workbook is opened, accepting the index of the worksheet.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");

// Make the third worksheet active
workBook.SetActiveTab(2);

workBook.SaveAs("setActiveTab.xlsx");
```

<center>
    ![Set Active Worksheet](https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-set-active-tab.png)
</center>

<hr>

## Deleting a Worksheet
The `RemoveWorksheet` method is utilized to eliminate a worksheet using its index or name.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("createNewWorkSheets.xlsx");

// Delete the second worksheet using index
workBook.RemoveWorkSheet(1);

// Also, deleting another by name
workBook.RemoveWorkSheet("workSheet2");

workBook.SaveAs("removeWorksheet.xlsx");
```

<center>
    ![Remove Worksheet](https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-remove-worksheet.png)
</center>

## Copying Worksheets
Worksheets can be duplicated within the same workbook or transferred to a different one using the `CopySheet` and `CopyTo` methods, respectively.

```cs
using IronXL;

WorkBook firstBook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkBook secondBook = WorkBook.Create();

// Access the default worksheet
WorkSheet workSheet = firstBook.DefaultWorkSheet;

// Copy inside the same workbook
workSheet.CopySheet("Copied Sheet");

// Transfer sheet to another workbook
workSheet.CopyTo(secondBook, "Copied Sheet");

firstBook.SaveAs("firstWorksheet.xlsx");
secondBook.SaveAs("secondWorksheet.xlsx");
```

<div style="display: flex; justify-content: space-between;">
    ![First Worksheet](https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-copy-worksheet-first.png)
    ![Second Worksheet](https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-copy-worksheet-second.png)
</div>