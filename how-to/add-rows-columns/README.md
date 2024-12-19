# How to Add or Remove Rows and Columns Using IronXL

***Based on <https://ironsoftware.com/how-to/add-rows-columns/>***


IronXL simplifies the process of adding or removing rows and columns in C# without relying on Office Interop. Here's how you can manage your spreadsheet data more efficiently.

### Getting Started with IronXL

---

## Example: Adding New Rows

To add new rows to your Excel spreadsheet, utilize the `InsertRow` and `InsertRows` methods. These allow you to specify the precise location for the new rows. Note that adding rows directly on a filtered row may disrupt the file, making an Excel repair necessary to view the data again.

```cs
using IronXL;

// Open an existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Add a single row before the second row
workSheet.InsertRow(1);

// Add three rows starting after the third row
workSheet.InsertRows(3, 3);

workBook.SaveAs("updatedRows.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/add-rows-columns/add-rows-columns-rows.png" alt="Add New Row" class="img-responsive add-shadow">
    </div>
</div>
<hr>

## Remove a Row Example

To delete a row from the spreadsheet, use the `GetRow` method to locate it, then apply the `RemoveRow` method to it. Be aware that removing the header row is not an option.

```cs
using IronXL;

// Open an existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Delete the fifth row
workSheet.GetRow(4).RemoveRow();

workBook.SaveAs("removedRow.xlsx");
```

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/add-rows-columns/add-rows-columns-remove-row.png" alt="Remove Row" class="img-responsive add-shadow">
    </div>
</div>
<hr>

## Example: Inserting New Columns

Add new column(s) to specific locations within your sheet using the `InsertColumn` and `InsertColumns` methods. Bear in mind, inserting columns within certain ranges might lead to conflicts, requiring an Excel file repair for proper functionality.

```cs
using IronXL;

// Open an existing spreadsheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Insert a single column at the start
workSheet.InsertColumn(0);

// Insert two columns after the second column
workSheet.InsertColumns(2, 2);

workBook.SaveAs("updatedColumns.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/add-rows-columns/add-rows-columns-columns.png" alt="Add New Column" class="img-responsive add-shadow">
    </div>
</div>

Keep in mind that row and column indices start at zero.