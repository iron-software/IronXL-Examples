# How to Select a Range in Excel with IronXL

***Based on <https://ironsoftware.com/how-to/select-range/>***


IronXL simplifies the process of selecting and manipulating ranges in an Excel worksheet, offering a powerful alternative to Office Interop.

<h3>Getting Started with IronXL</h3>

-------------------------------------

## Example of Selecting a Range

IronXL enables a variety of operations on selected ranges, including [sorting](https://ironsoftware.com/csharp/excel/how-to/sort-cells/), calculations, and aggregations. When executing methods that modify or relocate cell values, the targeted range, row, or column will be automatically updated. IronXL also allows the merging of multiple `IronXL.Ranges.Range` objects using the '+' operator.

### Selecting a Range

To specify a range from cell **A2** to **B8**, you would use:

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Extract the desired range from the worksheet
var range = workSheet["A2:B8"];
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/select-range/select-range-range.png" alt="Select Range" class="img-responsive add-shadow">
    </div>
</div>

### Selecting a Row

To select the fourth row, make use of the `GetRow(3)` method which relies on zero-based indexing. This function will ensure that the selected row captures all corresponding cells, regardless of whether they hold data or not.

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Access the fourth row
var row = workSheet.GetRow(3);
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/select-range/select-range-row.png" alt="Select Row" class="img-responsive add-shadow">
    </div>
</div>

### Selecting a Column

For selecting column C, use the `GetColumn(2)` method or simply reference the column as `workSheet["C:C"]`. Much like when selecting rows, this approach will also include all pertinent cells in the column.

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Retrieve column C
var column = workSheet.GetColumn(2);
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/select-range/select-range-column.png" alt="Select Column" class="img-responsive add-shadow">
    </div>
</div>

All indexing for rows and columns adhere to zero-based indexing.

### Combining Ranges

IronXL allows the merging of multiple `IronXL.Ranges.Range` instances utilizing the '+' operator. This feature offers a convenient way to extend or merge ranges, thereby creating a new comprehensive range. Note that using the '+' operator to connect rows and columns directly is not supported. 

The example below illustrates modifying an original range to incorporate additional combined ranges:

```cs
using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Access and modify the initial range
var range = workSheet["A2:B2"];

// Combine the initial range with another
var combinedRange = range + workSheet["A5:B5"];
```