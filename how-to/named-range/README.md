# How to Add Named Range

***Based on <https://ironsoftware.com/how-to/named-range/>***


A named range is a predefined block of cells given a unique label. Instead of referencing these cells through their cell coordinates (such as A1:B10), you can simply name them, which simplifies referencing them in formulas and calculations. For instance, by labeling a range as "SalesData," you can use `SUM(SalesData)` in a formula rather than directly specifying the cell addresses.

<h3>Get Started with IronXL</h3>

-------------------------------------

## Add Named Range Example

To create a named range, utilize the `AddNamedRange` method where you provide the name for the named range as a string and the corresponding range object.

```cs
using IronXL;

WorkBook workBook = WorkBook.Create();
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Define range
var selectedRange = workSheet["A1:A5"];

// Create named range
workSheet.AddNamedRange("range1", selectedRange);

workBook.SaveAs("addNamedRange.xlsx");
```

<div class="content-img-align-center">
    <div class="center-image-wrapper">
        <img src="https://ironsoftware.com/static-assets/excel/how-to/named-range/named-range.webp" alt="Named Range" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Retrieve Named Range Example

### Retrieve All Named Ranges

The `GetNamedRanges` method fetches a list containing all named ranges within the worksheet.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Obtain all named ranges
var namedRangeList = workSheet.GetNamedRanges();
```

### Retrieve Specific Named Range

Use the `FindNamedRange` method to obtain the exact reference for the named range specified, like Sheet1!$A$1:$A$5. This reference can subsequently be used both for referencing and selecting the relevant cell range. Be sure to include the worksheet name.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Retrieve named range reference
string namedRangeAddress = workSheet.FindNamedRange("range1");

// Access specific range
var range = workSheet[$"{namedRangeAddress}"];
```

<hr>

## Remove Named Range Example

To delete a named range, employ the `RemoveNamedRange` method with the name of the named range specified as a string.

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("addNamedRange.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;

// Delete named range
workSheet.RemoveNamedRange("range1");
```