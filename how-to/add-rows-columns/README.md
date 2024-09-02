# How to Add Rows and Columns Using IronXL

The IronXL library simplifies the process of adding rows and columns in C# without the need for Office Interop. This guide provides clear instructions on how to expand your spreadsheets effectively.

## Example: Adding New Rows

The IronXL library facilitates adding new rows through the `InsertRow` and `InsertRows` methods, which allow insertion at a specified index.

It's important to avoid inserting rows directly on the filter row, as this may lead to conflicts and necessitate the use of Excel repair tools to resolve spreadsheet display issues.

```cs
using IronXL;

// Load an existing spreadsheet
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Insert a single row before the 2nd row
worksheet.InsertRow(1);

// Insert three rows starting after the 3rd row
worksheet.InsertRows(3, 3);

workbook.SaveAs("addRow.xlsx");
```

![Add New Row](https://ironsoftware.com/static-assets/excel/how-to/add-rows-columns/add-rows-columns-rows.png)
<hr>

## Example: Removing a Row

To remove a specific row, use the `GetRow` method to locate it and `RemoveRow` to delete it.

Note that the header row cannot be removed.

```cs
using IronXL;

// Open an existing spreadsheet
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Select and remove the 5th row
worksheet.GetRow(4).RemoveRow();

workbook.SaveAs("removeRow.xlsx");
```

![Remove a Row](https://ironsoftware.com/static-assets/excel/how-to/add-rows-columns/add-rows-columns-remove-row.png)
<hr>

## Example: Adding New Columns

Use the `InsertColumn` and `InsertColumns` methods to add new columns before a specified index in your spreadsheet.

Inserting columns within occupied table ranges may cause issues requiring Excel repair to correct the file. If attempting to add a column to an entirely empty sheet, a `System.InvalidOperationException` can occur, with the message 'Sequence contains no elements'.

To clean up any extraneous empty rows and columns along the borders of your data range, employ the [Trim() method](https://ironsoftware.com/csharp/excel/how-to/trim-cell-range/). However, direct removal of columns is not supported.

```cs
using IronXL;

// Open an existing spreadsheet
WorkBook workbook = WorkBook.Load("sample.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Insert a column before the first column
worksheet.InsertColumn(0);

// Insert two columns after the second column
worksheet.InsertColumns(2, 2);

workbook.SaveAs("addColumn.xlsx");
```

![Add New Column](https://ironsoftware.com/static-assets/excel/how-to/add-rows-columns/add-rows-columns-columns.png)

All indices for rows and columns are based on zero indexing for clarity and consistency.