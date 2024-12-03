# Implementing Freeze Panes in Excel

***Based on <https://ironsoftware.com/how-to/add-freeze-panes/>***


## Overview

Navigating through extensive spreadsheets with over **50 rows** or spanning beyond column **'Z'** while maintaining header visibility can be challenging. Utilizing the **Freeze Pane** feature addresses this issue effectively.

## Implementing Freeze Panes in Spreadsheets

The freeze panes feature allows users to lock specific rows and columns, keeping them visible as you navigate through the spreadsheet. This functionality is particularly useful for maintaining the visibility of headers while comparing data across various parts.

### FreezePane(int column, int row)

The `FreezePane` method facilitates freezing panes by specifying the starting column and row. The selected column and row will not be included in the frozen section. For instance, invoking `workSheet.FreezePane(1, 4)` will freeze the panes up to **column A** and **rows 1 to 3**.

Here’s how to implement a freeze pane covering columns A to B and rows 1 to 3:

```cs
using System.Linq;
using IronXL.Excel;
namespace ironxl.AddFreezePanes
{
    public class FreezePanesExample
    {
        public void Execute()
        {
            WorkBook workbook = WorkBook.Load("sample.xlsx");
            WorkSheet worksheet = workbook.WorkSheets.First();
            
            // Locking the rows and columns: A-B and 1-3
            worksheet.FreezePane(2, 3);
            
            workbook.SaveAs("enhancedFreezePanes.xlsx");
        }
    }
}
```

### Visualization

![Freeze Panes in Action](https://www.ironsoftware.com/static-assets/excel/how-to/add-freeze-panes/add-freeze-panes-add.gif)

## Removing Freeze Panes

The `RemovePane` function is used to effortlessly clear all freeze panes configurations.

```cs
using IronXL.Excel;
namespace ironxl.AddFreezePanes
{
    public class RemovalOfPanes
    {
        public void Execute()
        {
            // Clearing any freeze or split panes
            worksheet.RemovePane();
        }
    }
}
```

## Advanced Freeze Pane Implementation

The `FreezePane` offers enhanced functionality, including pre-scrolling settings.

### FreezePane(int column, int row, int subsequentColumn, int subsequentRow)

This method allows freezing specific areas, as well as setting up a predefined scroll within the rows and columns, enhancing initial visibility upon opening the document. For example, `workSheet.FreezePane(5, 2, 6, 7)` will initially display **columns A-E, columns from G onwards**, and **rows 1-2, rows from 8 onwards**.

Freeze pane setups will replace any existing configurations, and are not compatible with the older .xls format used in MS Excel versions 97-2003.

```cs
using System.Linq;
using IronXL.Excel;
namespace ironxl.AddFreezePanes
{
    public class AdvancedFreezePanes
    {
        public void Execute()
        {
            WorkBook workbook = WorkBook.Load("sample.xlsx");
            WorkSheet worksheet = workbook.WorkSheets.First();
            
            // Setting up an advanced freeze pane with pre-scroll
            worksheet.FreezePane(5, 5, 6, 7);
            
            workbook.SaveAs("advancedFreezePanes.xlsx");
        }
    }
}
```

### Showcase

<div style="text-align:center">
    <img src="https://www.ironsoftware.com/static-assets/excel/how-to/add-freeze-panes/add-freeze-panes-advance.png" alt="Advanced Freeze Panes Demonstration" style="margin:auto; display:block; max-width:100%; height:auto;"/>
</div>