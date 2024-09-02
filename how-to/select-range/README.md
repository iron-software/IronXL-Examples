# How to Select Ranges in Excel with IronXL

IronXL simplifies the process of selecting and manipulating ranges in Excel spreadsheets without the need for Office Interop.

## Example of Selecting a Range

IronXL empowers you to carry out numerous functions with selected ranges such as [sorting](https://ironsoftware.com/csharp/excel/how-to/sort-cells/), computations, and summary operations. When you use methods that change or shuffle cell values, the specified range, row, or column adjusts its values automatically. IronXL also supports the merging of multiple `IronXL.Ranges.Range` objects using the `+` operator.

### Selecting a Specific Range

For selecting a range from cell **A2** to **B8**, the code snippet below demonstrates how to perform this action:

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Select a specific range from the worksheet
var range = workSheet["A2:B8"];
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/select-range/select-range-range.png" alt="Select Range" class="img-responsive add-shadow">
    </div>
</div>

### Selecting a Row

To select the 4th row, utilize the `GetRow(3)` function with zero-based indexing. This includes all cells across the specified row, accounting for any cells that may be populated in the same column in other rows.

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Obtain the fourth row from the worksheet
var row = workSheet.GetRow(3);
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/select-range/select-range-row.png" alt="Select Row" class="img-responsive add-shadow">
    </div>
</div>

### Selecting a Column

Selecting column C in your spreadsheet can be done using `GetColumn(2)` or by specifying the range as `workSheet["C:C"]`. This method also ensures all cells in the column are selected, irrespective of their populated state in other rows.

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Access the third column from the worksheet
var column = workSheet.GetColumn(2);
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/select-range/select-range-column.png" alt="Select Column" class="img-responsive add-shadow">
    </div>
</div>

All indexing for rows and columns follow a zero-based notation.

### Combining Ranges

IronXL allows you to combine multiple `IronXL.Ranges.Range` objects using the `+` operator, facilitating easy concatenation or merging of ranges, resulting in a new, extended range.

It's important to note that combining rows and columns directly via the `+` operator is unsupported.

The code below modifies the `range` variable to encompass the combined ranges after merging them.

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Initialize range selection
var range = workSheet["A2:B2"];

// Combine with another range
var combinedRange = range + workSheet["A5:B5"];
```