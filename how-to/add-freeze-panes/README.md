# How to Implement Freeze Pane in Excel

***Based on <https://ironsoftware.com/how-to/add-freeze-panes/>***


## Introduction

In extensive spreadsheets with over **50 rows** or beyond the **'Z' column**, maintaining a view of the relevant headers while scrolling through data can be problematic. The **Freeze Pane** feature elegantly addresses this by keeping rows or columns stationary.

<h3>Begin with IronXL</h3>

---

## Implementing Freeze Pane

The freeze pane feature allows for certain rows and columns to be fixed, making them perpetually visible while navigating through the spreadsheet. This is especially beneficial for keeping the header row or column static as you assess other detailed data.

### CreateFreezePane(int column, int row)

The `CreateFreezePane` method is used to establish a freeze pane, with inputs defining the beginning column and row. The initial column and row are excluded from being frozen. For example, `workSheet.CreateFreezePane(1, 4)` sets up a freeze pane commencing at **column A** and **rows 1-4**.

Below is the code snippet that illustrates initiating a freeze pane covering columns A to B and rows 1 to 3:

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Establishing freeze pane at columns A-B and rows 1-3
workSheet.CreateFreezePane(2, 3);

workBook.SaveAs("createFreezePanes.xlsx");
```

### Visualization

<img src="https://ironsoftware.com/static-assets/excel/how-to/add-freeze-panes/add-freeze-panes-add.gif" alt="Freeze Pane in Action" class="img-responsive add-shadow" style="margin-bottom: 30px;"/>

## Removing Freeze Pane

To undo the freeze pane, utilize the `RemovePane` method that eliminates any existing freeze panes in the worksheet.

```cs
// Eliminating any freeze or split pane
workSheet.RemovePane();
```

## Advanced Freeze Pane Setup

`CreateFreezePane` can also be employed for more sophisticated freezing that includes scrolling capabilities.

### CreateFreezePane(int column, int row, int subsequentColumn, int subsequentRow)
The method not only allows for freezing specific rows and columns but also adds.scroll functionality within the sheet.

For instance, utilizing `workSheet.CreateFreezePane(5, 2, 6, 7)` establishes a freeze pane over **columns A-E** and **rows 1-2**, and includes a scroll feature that will display **columns A-E, G-...** and **rows 1-2, 8-...** upon opening.

A single freeze pane setting is used at a time; any new settings will overwrite existing ones.

Note: Freeze panes are not compatible with Microsoft Excel versions 97-2003 (.xls).

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Applying advanced freeze pane at column A-E and row 1-5; prescroll to columns E, G,... and rows 5, 8,...
workSheet.CreateFreezePane(5, 5, 6, 7);

workBook.SaveAs("createAdvancedFreezePanes.xlsx");
```

### Advanced Demonstration
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/add-freeze-panes/add-freeze-panes-advance.png" alt="Advanced Freeze Panes Demonstration" class="img-responsive add-shadow">
    </div>
</div>