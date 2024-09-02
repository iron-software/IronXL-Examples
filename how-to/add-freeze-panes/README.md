# How to Implement Freeze Panes in Excel Files

## Overview

Working with substantial spreadsheets that contain tens of rows or span multiple columns past `Z` can be cumbersome, especially when needing to reference header rows or columns. Freeze Panes functionality serves as an effective solution to keep these headers visible while you scroll through your data.

## Implementing Freeze Panes

The Freeze Panes feature allows you to lock specific rows and columns so that they stay visible as you scroll through your spreadsheet. This can greatly enhance the usability of your worksheet by keeping headers fixed while you review or compare other rows or columns of data.

### Usage of `CreateFreezePane(int column, int row)`

To implement a freeze pane, the `CreateFreezePane` method is used, where you specify the starting column and row. These starting points are excluded from the frozen pane. For instance, calling `workSheet.CreateFreezePane(1, 4)` would freeze from **column A** to **the first three rows**.

Here's how to create a freeze pane across columns A to B for the first three rows:

```cs
// Necessary namespaces for using IronXL functions
using IronXL;

// Load workbook and select first worksheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Applying freeze pane from column A to B, across the first three rows
workSheet.CreateFreezePane(2, 3);

// Save changes to a new file
workBook.SaveAs("createFreezePanes.xlsx");
```

### Visual Guide

![Freeze Pane in Action](https://ironsoftware.com/static-assets/excel/how-to/add-freeze-panes/add-freeze-panes-add.gif)

## Removing Freeze Panes

To remove any established freeze panes, simply use the `RemovePane` method.

```cs
// This command removes any freeze or split pane applied to the worksheet
workSheet.RemovePane();
```

## Advanced Freeze Pane Configuration

The `CreateFreezePane` method also allows for more sophisticated setups, including predefined scrolling.

### `CreateFreezePane(int column, int row, int subsequentColumn, int subsequentRow)`

This variation not only sets where the freeze starts but also defines scrolling behavior for the sheet. For example, `workSheet.CreateFreezePane(5, 2, 6, 7)` creates a freeze across **columns A-E** and **rows 1-2** and will position the view to include **columns G and onwards** and **rows 8 and onwards** upon opening the sheet.

Note that you can only set one freeze pane at a time; setting a new one overrides the previous settings.

```cs
// Namespace inclusion for IronXL functions
using IronXL;

// Load the workbook and select the first sheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Setting advanced freeze pane with prescrolling
workSheet.CreateFreezePane(5, 5, 6, 7);

// Save the new setup to a file
workBook.SaveAs("createFreezePanes.xlsx");
```

### Advanced Demonstration
<div style="text-align: center;">
  <img src="https://ironsoftware.com/static-assets/excel/how-to/add-freeze-panes/add-freeze-panes-advance.png" alt="Advanced Freeze Panes Demonstration" style="width: auto; max-width: 100%; height: auto; margin-bottom: 30px;">
</div>

**Note:** The Freeze Pane functionality is not compatible with Microsoft Excel versions 97-2003 (.xls).