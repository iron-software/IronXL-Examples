# Sorting Cell Ranges

***Based on <https://ironsoftware.com/how-to/sort-cells/>***


Sorting Excel data by values or alphabetically is critical for thorough data analysis. IronXL simplifies the process of sorting columns, rows, and cell ranges in both C# and VB.NET environments.

### Getting Started with IronXL

---

## Example: Sorting Columns Independently

Apply sorting to selected ranges or columns using the `SortAscending` or `SortDescending` methods as per the required order.

When sorting across a range with multiple columns, the `SortAscending` or `SortDescending` methods operate on each column individually.

These methods position any empty cells at either the top or bottom of the range. To clean up these empty cells, use the [Trim method](https://ironsoftware.com/csharp/excel/how-to/trim-cell-range/) post-sorting, which helps in maintaining a tidy dataset.

```cs
using IronXL;

// Load an Excel workbook
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Retrieve the first column
var column = workSheet.GetColumn(0);

// Sort the first column in ascending order (A to Z)
column.SortAscending();

// Sort the first column in descending order (Z to A)
column.SortDescending();

// Save the sorted range to a new file
workBook.SaveAs("sortExcelRange.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/sort-cells/sort-cells-range.png" alt="Sort Ascending and Descending" class="img-responsive add-shadow">
    </div>
</div>

---

## Example: Sorting by a Specific Column

To sort a range based on a specific column, the `SortByColumn` method is utilized requiring two parameters: the column identifier for sorting, and the range over which the sorting should be applied.

```cs
using IronXL;

// Load a workbook
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Define the range to be sorted
var range = workSheet["A1:D10"];

// Sort the defined range by the second column in ascending order
range.SortByColumn("B", SortOrder.Ascending);

// Save the workbook with the sorted range
workBook.SaveAs("sortRange.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/sort-cells/sort-cells-sort-by-column.png" alt="Sort by Specific Column" class="img-responsive add-shadow">
    </div>
</div>

Currently, sorting by more than one column sequentially (e.g., first by column A then by column B) is not supported.