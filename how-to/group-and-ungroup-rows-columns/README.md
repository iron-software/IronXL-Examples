# How to Segment and Revert Segmentation for Excel Rows & Columns

## Overview

Excel's grouping functionality allows users to organize and manage data by creating expandable and collapsible sections for rows or columns. This is particularly helpful for managing large datasets by enabling focused analysis on specific sections. In contrast, the ungrouping function reverts data to its original unpartitioned format, improving clarity and data presentation.

IronXL supports C# .NET developers in programmatically implementing these tasks without needing Interop services.

## Example: Grouping and Ungrouping Excel Rows

For all the examples given, we use zero-based indexing, and note that these operations can only be applied to cells that contain data.

### Grouping Rows

The `GroupRows` method allows the setting of row grouping by specifying the start and end indexes. This method can be reused for multiple or different rows to apply numerous groupings.

```cs
using IronXL;

// Open an existing file
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Group rows from index 0 to index 7
worksheet.GroupRows(0, 7);

workbook.SaveAs("groupedRows.xlsx");
```

#### Visualization

![Group Rows](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-group-row.png)

### Ungrouping Rows

Undo the grouping of rows using the `UngroupRows` method. Apply this method to segment a grouped range into two, but do note the resulting segments aren't autonomous groups.

```cs
using IronXL;

// Open an existing file
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Ungroup rows from index 2 to 4
worksheet.UngroupRows(2, 4);

workbook.SaveAs("ungroupedRows.xlsx");
```

#### Results Display

<div class="competitors-section__wrapper-even-1">
    <div class="competitors__card" style="width: 100%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-group-row.png" alt="Group Rows" class="img-responsive add-shadow" style="margin-bottom: 20px;"/>
        <p class="competitors__download-link" style="color: #181818; font-style: italic; margin-bottom: 30px;">Before</p>
    </div>
    <div class="competitors__card" style="width: 100%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-ungroup-row.png" alt="Ungroup Rows" class="img-responsive add-shadow" style="margin-bottom: 20px;"/>
        <p class="competitors__download-link" style="color: #181818; font-style: italic; margin-bottom: 30px;">After</p>
    </div>
</div>

## Grouping and Ungrouping Columns

### Grouping Columns

Columns are grouped similarly to rows. Use the `GroupColumns` method by specifying either the index or the alphabetical reference of the columns.

```cs
using IronXL;

// Open an existing file
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Group columns from A to F
worksheet.GroupColumns(0, 5);

workbook.SaveAs("groupedColumns.xlsx");
```

#### Visualization

![Group Columns](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-group-column.png)

### Ungrouping Columns

Ungrouping columns works on indexing or alphabetical reference, utilizing either `UngroupColumn` or `UngroupColumns`.

```cs
using IronXL;

// Open an existing spreadsheet file
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Ungroup columns C to D
worksheet.UngroupColumn("C", "D");

workbook.SaveAs("ungroupedColumns.xlsx");
```

#### Results Display

<div class="competitors-section__wrapper-even-1">
    <div class="competitors__card" style="width: 100%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-group-column.png" alt="Group Columns" class="img-responsive add-shadow" style="margin-bottom: 20px;"/>
        <p class="competitors__download-link" style="color: #181818; font-style: italic; margin-bottom: 30px;">Before</p>
    </div>
    <div class="competitors__card" style="width: 100%;">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-ungroup-column.png" alt="Ungroup Columns" class="img-responsive add-shadow" style="margin-bottom: 20px;"/>
        <p class="competitors__download-link" style="color: #181818; font-style: italic;">After</p>
    </div>
</div>