# How to Add New Rows and Columns

***Based on <https://ironsoftware.com/how-to/add-rows-columns/>***


The IronXL library simplifies the addition of rows and columns in C# without the need for Office Interop.


## Example on How to Add a New Row

To add new rows to a spreadsheet, utilize the `InsertRow` and `InsertRows` methods. These functions allow you to specify the exact position for new rows.

However, be cautious when inserting rows at the filter row position as this can lead to conflicts, potentially requiring Excel repair operations for proper file handling.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.AddRowsColumns
{
    public class Section1
    {
        public void Run()
        {
            // Opening an existing spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Inserting a single row at index 1
            workSheet.InsertRow(1);
            
            // Adding three rows starting from index 3
            workSheet.InsertRows(3, 3);
            
            // Save the updates to a new file
            workBook.SaveAs("addRow.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/add-rows-columns/add-rows-columns-rows.png" alt="Add New Row" class="img-responsive add-shadow">
    </div>
</div>
<hr>

## Example on How to Remove a Row

To delete a row, make use of the `GetRow` method to select it and `RemoveRow` to delete it from your sheet.

Note that it's not possible to delete the spreadsheet's header row.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.AddRowsColumns
{
    public class Section2
    {
        public void Run()
        {
            // Load the spreadsheet
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Target and remove the fifth row
            workSheet.GetRow(4).RemoveRow();
            
            // Saving the file after changes
            workBook.SaveAs("removeRow.xlsx");
        }
    }
}
```

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/add-rows-columns/add-rows-columns-remove-row.png" alt="Add New Row" class="img-responsive add-shadow">
    </div>
</div>
<hr>

## How to Insert a New Column Example

You can add new columns to your spreadsheet by employing the `InsertColumn` and `InsertColumns` methods, targeting specific index positions.

Be mindful that inserting columns within the range of the table might cause file conflicts, potentially requiring Excel repair to correct.

For managing space efficiency, use the [`Trim()`](https://ironsoftware.com/csharp/excel/how-to/trim-cell-range/) technique to eliminate extraneous rows and columns along the boundaries of the range. Current operations do not allow for the removal of columns, and adding a column to an empty sheet might result in a `System.InvalidOperationException`.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.AddRowsColumns
{
    public class Section3
    {
        public void Run()
        {
            // Open the spreadsheet file
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            WorkSheet workSheet = workBook.DefaultWorkSheet;
            
            // Adding a column at the beginning
            workSheet.InsertColumn(0);
            
            // Adding two columns starting from the third column
            workSheet.InsertColumns(2, 2);
            
            // Save the modified file
            workBook.SaveAs("addColumn.xlsx");
        }
    }
}
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/add-rows-columns/add-rows-columns-columns.png" alt="Add New Column" class="img-responsive add-shadow">
    </div>
</div>

Index positions for rows and columns start from zero.