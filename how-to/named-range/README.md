# How to Create and Manage Named Ranges in Excel Documents

A named range in Excel designates a group of cells within the spreadsheet that you identify with a distinct label. This unique name allows you to refer easily to those cells in formulas—instead of using traditional cell references like A1:B10, you might use a meaningful name such as "SalesData." This approach simplifies formula construction and increases the readability of your functions, letting you use notation like SUM(SalesData) instead of the cell range directly.

## Example of Adding a Named Range

To create a named range, utilize the `AddNamedRange` method. Here, you pass in a string that will serve as the named range's identifier along with the range object itself.

```cs
using IronXL;

// Create a new workbook and select the default worksheet
WorkBook workbook = WorkBook.Create();
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Define a range of cells
var rangeToName = worksheet["A1:A5"];

// Assign the selected range a name
worksheet.AddNamedRange("range1", rangeToName);

// Save the workbook with the new named range
workbook.SaveAs("addNamedRange.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/named-range/named-range.webp" alt="Named Range" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Example of Retrieving Named Ranges

### Retrieve All Named Ranges

Retrieve a list of all named ranges in the worksheet with the `GetNamedRanges` method.

```cs
using IronXL;

// Load an existing workbook and select the default worksheet
WorkBook workbook = WorkBook.Load("addNamedRange.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Fetch all named ranges
var allNamedRanges = worksheet.GetNamedRanges();
```

### Retrieve a Specific Named Range

Find a specific named range and get its complete reference using the `FindNamedRange` method.

```cs
using IronXL;

// Load the workbook and select the default worksheet
WorkBook workbook = WorkBook.Load("addNamedRange.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Find the absolute reference of a specific named range
string namedRangeReference = worksheet.FindNamedRange("range1");

// Use the reference to select the named range
var selectedRange = worksheet[$"{namedRangeReference}"];
```

<hr>

## Example of Removing a Named Range

Eliminate a named range from the worksheet by using the `RemoveNamedRange` method.

```cs
using IronXL;

// Open the workbook where the named range exists
WorkBook workbook = WorkBook.Load("addNamedRange.xlsx");
WorkSheet worksheet = workbook.DefaultWorkSheet;

// Remove a specific named range
worksheet.RemoveNamedRange("range1");
```