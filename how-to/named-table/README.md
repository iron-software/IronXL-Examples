# How to Insert a Named Table

***Based on <https://ironsoftware.com/how-to/named-table/>***


A named table, often referred to as an Excel Table, is a range that has been distinctly named and possesses enhanced features and capabilities.

### Introduction to IronXL

---

## Example of Adding a Named Table

To insert a named table within a spreadsheet, utilize the `AddNamedTable` method. This method requires the table's name as a string, a range object, and optionally, you can set the table style and enable a filter view.

```cs
using IronXL;
using IronXL.Styles;

// Create a new workbook and access the default worksheet
WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Populate data in a range 
workSheet["A2:C5"].StringValue = "Sample Text";

// Define the range for the named table
var selectedRange = workSheet["A1:C5"];
bool enableFilter = false;
var tableStyle = TableStyle.TableStyleDark1;

// Creating the named table
workSheet.AddNamedTable("table1", selectedRange, enableFilter, tableStyle);

// Save workbook to file
workBook.SaveAs("addNamedTable.xlsx");
```

<div  class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/named-table/named-table.webp" alt="Named Table" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Example of Retrieving Named Tables

### Fetch All Named Tables

Retrieve all the named tables present in the worksheet using the `GetNamedTableNames` method, which returns a list of their names.

```cs
using IronXL;

// Load the workbook and get the default worksheet
WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Fetch all named tables
var namedTableList = workSheet.GetNamedTableNames();
```

### Access a Specific Named Table

To get a particular named table from the worksheet, use the `GetNamedTable` method.

```cs
using IronXL;

// Load the workbook and access the default worksheet
WorkBook workBook = WorkBook.Load("addNamedTable.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Retrieve a specific named table
var namedRangeAddress = workSheet.GetNamedTable("table1");
```

IronXL also supports creating named ranges. Discover more at [How to Add Named Range](https://ironsoftware.com/csharp/excel/how-to/named-range/).