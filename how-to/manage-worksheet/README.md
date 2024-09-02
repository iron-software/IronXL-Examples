# Managing Worksheets with IronXL

The **IronXL** library provides efficient management of worksheets in Excel through straightforward C# integration. Users can create, delete, reposition, and set active worksheets in Excel files effortlessly, bypassing the need for Office Interop compatibility.

## Example of Worksheet Management

IronXL facilitates essential management tasks for worksheets via intuitive, single-statement commands.

Index positions used throughout this document are zero-based.

## Creating a Worksheet

The `CreateWorksheet` method makes adding a new worksheet simple, where you only need to specify the desired worksheet's name. Furthermore, this method returns a worksheet object, which can be immediately used for further operations such as [merging cells](https://ironsoftware.com/csharp/excel/how-to/csharp-excel-merge-cells/).

```cs
using IronXL;

// Initialize a new Excel workbook
WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);

// Adding multiple worksheets
WorkSheet ws1 = workbook.CreateWorkSheet("SheetOne");
WorkSheet ws2 = workbook.CreateWorkSheet("SheetTwo");
WorkSheet ws3 = workbook.CreateWorkSheet("SheetThree");
WorkSheet ws4 = workbook.CreateWorkSheet("SheetFour");

workbook.SaveAs("NewSheets.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-create-worksheet.png" alt="Create Worksheets" class="img-responsive add-shadow">
    </div>
</div>

## Changing Worksheet Position

You can relocate a worksheet within a workbook using `SetSheetPosition`, which requires the worksheet name and the new index position.

```cs
using IronXL;

WorkBook workbook = WorkBook.Load("NewSheets.xlsx");

// Adjusting the position of a worksheet
workbook.SetSheetPosition("SheetTwo", 0);

workbook.SaveAs("AdjustedWorksheetPosition.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-set-sheet-position.png" alt="Change Worksheet Position" class="img-responsive add-shadow">
    </div>
</div>

## Activating a Worksheet

Set which worksheet opens by default with `SetActiveTab`, using the worksheet’s index position.

```cs
using IronXL;

WorkBook workbook = WorkBook.Load("NewSheets.xlsx");

// Activate a specific worksheet
workbook.SetActiveTab(2);

workbook.SaveAs("ActiveWorksheet.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-set-active-tab.png" alt="Set Active Worksheet" class="img-responsive add-shadow">
    </div>
</div>

## Removing a Worksheet

Worksheets can be deleted by specifying either their name or index position using the `RemoveWorksheet` method.

```cs
using IronXL;

WorkBook workbook = WorkBook.Load("NewSheets.xlsx");

// Deleting worksheets
workbook.RemoveWorkSheet(1);
workbook.RemoveWorkSheet("SheetTwo");

workbook.SaveAs("WorksheetRemoved.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-remove-worksheet.png" alt="Remove Worksheet" class="img-responsive add-shadow">
    </div>
</div>

## Copying a Worksheet

Worksheets can be copied to the same or different workbooks using `CopySheet` and `CopyTo` respectively.

```cs
using IronXL;

WorkBook originalBook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkBook targetBook = WorkBook.Create();

// Selecting the default worksheet
WorkSheet defaultSheet = originalBook.DefaultWorkSheet;

// Copy within the same workbook
defaultSheet.CopySheet("DuplicateSheet");

// Copy to another workbook
defaultSheet.CopyTo(targetBook, "DuplicateSheet");

originalBook.SaveAs("OriginalWorkbook.xlsx");
targetBook.SaveAs("TargetWorkbook.xlsx");
```

<div class="competitors-section__wrapper-even-1">
    <div class="competitors__card" style="width: 54%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-copy-worksheet-first.png" alt="First Worksheet" class="img-responsive add-shadow" style="margin-bottom: 20px;"/>
        <p class="competitors__download-link" style="color: #181818; font-style: italic; margin-bottom: 30px;">OriginalWorkbook.xlsx</p>
    </div>
    <div class="competitors__card" style="width: 44%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/manage-worksheet/manage-worksheet-copy-worksheet-second.png" alt="Second Worksheet" class="img-responsive add-shadow" style="margin-bottom: 20px;"/>
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">TargetWorkbook.xlsx</p>
    </div>
</div>