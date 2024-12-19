# Row and Column Grouping in Excel with IronXL

***Based on <https://ironsoftware.com/how-to/group-and-ungroup-rows-columns/>***


## Overview

In Excel, data management is enhanced through the use of grouping features, which allow users to create collapsible sections within rows or columns, making large datasets easier to handle and analyze. Ungrouping, on the other hand, reverses these settings to display the spreadsheet in its full detail once more. These functionalities are essential for making large amounts of data more manageable for detailed review.

Utilizing IronXL, developers can implement such grouping and ungrouping features programmatically in C# .NET applications, all without depending on Office Interop.


<h3>Getting Started with IronXL</h3>
----------------------------------

## Example of Grouping and Ungrouping Rows

Keep in mind that the indices used here are zero-based.

### Grouping Rows

To create row groups in spreadsheets, the `WorkSheet.GroupRows` method is employed, which requires the start and end indices as parameters. It’s possible to execute multiple calls to this method on different or the same rows for additional groupings.

```cs
using IronXL;

// Initialize and load a workbook
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Group rows from the first to the eighth
worksheet.GroupRows(0, 7);

workbook.SaveAs("groupedRows.xlsx");
```

#### Grouped Rows Visual

![View of Grouped Rows](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-group-row.png "Group Rows Example")

### Ungrouping Rows

The `WorkSheet.UngroupRows` method serves to split grouped rows, essentially 'cutting' through them to segment the group effectively. It’s worth noting that despite the division, the newly separated areas will not be considered as new groups. For example, cutting through rows 3-5 that belong to a group from rows 0-8 will produce two groups: 0-2 and 6-8.

```cs
using IronXL;

// Initialize and load a workbook
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Ungroup rows 3 through 5
worksheet.UngroupRows(2, 4);

workbook.SaveAs("ungroupedRows.xlsx");
```

#### Ungrouped Rows Visual

![Before and After Ungrouping Rows](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-ungroup-row.png "Ungroup Rows Example")

## Example of Grouping and Ungrouping Columns

### Grouping Columns

Columns can be grouped similarly to rows using the `WorkSheet.GroupColumns` method. Specifying either the start and end indices or the alphabetical column identifiers can indicate the range to group.

```cs
using IronXL;

// Load the workbook
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Group columns from A to F
worksheet.GroupColumns(0, 5);

workbook.SaveAs("groupedColumns.xlsx");
```

#### Grouped Columns Visual

![Grouped Columns Example](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-group-column.png "Visual Representation of Grouped Columns")

### Ungrouping Columns

To ungroup columns that have been grouped, utilize either the `WorkSheet.UngroupColumn` method for specific columns by their alphabetical identifiers or `WorkSheet.UngroupColumns` by indices. This function splits the group, resulting in two separate sections.

```cs
using IronXL;

// Initialize and load the workbook
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Ungroup columns C to D
worksheet.UngroupColumn("C", "D");

workbook.SaveAs("ungroupedColumns.xlsx");
```

#### Ungrouped Columns Visual

![Before and After Ungrouping Columns](https://ironsoftware.com/static-assets/excel/how-to/group-and-ungroup-rows-columns/group-and-ungroup-rows-columns-ungroup-column.png "Ungroup Columns Example")