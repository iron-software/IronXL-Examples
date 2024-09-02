# How to Add a Named Table in Excel

A named table, often referred to as an Excel Table, is a designated range within an Excel sheet that is given a unique name and comes with enhanced features and capabilities.

## Example of Adding a Named Table

To insert a named table, utilize the `AddNamedTable` method. This method necessitates specifying the table's name as a string, the range to be converted into the table, and optionally, the table style and a filter toggle.

```cs
using IronXL;
using IronXL.Styles;

// Initialize a new workbook and select the default worksheet
WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Populate data in the worksheet
workSheet["A2:C5"].StringValue = "Data Entry";

// Define range and parameters for the named table
var tableRange = workSheet["A1:C5"];
bool enableFilter = false;
var styleOfTable = TableStyle.TableStyleDark1;

// Create the named table in the worksheet
workSheet.AddNamedTable("MyTable", tableRange, enableFilter, styleOfTable);

// Save the workbook to a file
workBook.SaveAs("addNamedTable.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://www.ironsoftware.com/static-assets/excel/how-to/named-table/named-table.webp" alt="Named Table" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Retrieving Named Tables Example

### Retrieve All Named Tables

The `GetNamedTableNames` method fetches all the named tables within a provided worksheet and returns them as a list of string names.

```cs
using IronXL;

// Load the workbook and select the default worksheet
WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Retrieve a list of all named table names in the worksheet
var tableNames = workSheet.GetNamedTableNames();
```

### Retrieve a Specific Named Table

To access a particular named table directly, use the `GetNamedTable` method.

```cs
using IronXL;

// Load the workbook and access the default worksheet
WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Retrieve a specific named table by its name
var specificTable = workSheet.GetNamedTable("MyTable");
```

IronXL also supports adding named ranges. For further details, please visit [How to Add Named Range](https://www.ironsoftware.com/csharp/excel/how-to/named-range/).