***Based on <https://ironsoftware.com/how-to/c-sharp-excel-interop/>***

```cs
static void Main(string [] args)
{
    // Load Workbook
    WorkBook wb = WorkBook.Load("sample.xlsx");
    // Select a Worksheet
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Update cell A3's value
    ws ["A3"].Value = "Updated Value for A3";
    // Save changes to the same file
    wb.SaveAs("sample.xlsx");
}
```

This script will replace the existing content of cell `A3` with `Updated Value for A3`.

Additionally, multiple cells can be updated with a singular value using a range:

```cs
 ws ["A3:C3"].Value = "Unified New Value";
```

This will modify row 3 from cell `A3` to `C3`, replacing all values with `Unified New Value`.

Learn more about the [Range Function in C#](https://ironsoftware.com/csharp/excel/#excel-ranges) from these practical examples.

### Replacing Cell Values

The flexibility of IronXL allows for straightforward substitution of old values with new ones across different levels of granularity within an Excel file. This includes:

* Entire Worksheet
```cs
 WorkSheet.Replace("existing value", "replacement value");
 ```
* Specified rows
```cs
WorkSheet.Rows [RowIndex].Replace("current value", "new value");
``` 
* Specified columns
```cs
WorkSheet.Columns [ColumnIndex].Replace("old value", "new value");
```
* Defined range
```cs
WorkSheet ["RangeStart:RangeEnd"].Replace("old", "new");
```

To illustrate substituting values in a specified range, consider this example:

```cs
/**
Replace Cell Value Range
anchor-replace-cell-values
**/
using IronXL;
static void Main(string [] args)
{
    // Load Workbook
    WorkBook wb = WorkBook.Load("sample.xlsx");
    // Select a Worksheet
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Define a range and replace values from B5 to G5
    ws ["B5:G5"].Replace("Normal", "Improved");
    // Save changes to the same file
    wb.SaveAs("sample.xlsx");
}
```

This code alters values within the range `B5` to `G5` from `Normal` to `Improved`. You can find more insights about [Editing Excel Cell Values in a Range](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/#sample-edit-cell-values-in-range) on the IronXL web page.

### Removing Rows from an Excel File

Occasionally, it's necessary to remove entire rows from an Excel file while developing applications. This can be accomplished using the `RemoveRow()` function. Here's an example:

```cs
/**
Remove Row
anchor-remove-rows-of-excel-file
**/
using IronXL;
static void Main(string [] args)
{ 
    // Load Workbook
    WorkBook wb = WorkBook.Load("sample.xlsx");
    // Get specific Worksheet
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Remove row number 2
    ws.Rows [2].RemoveRow();
    // Save changes to the existing file
    wb.SaveAs("sample.xlsx");
}
```

This code removes row number `2` from the file named `sample.xlsx`.

<hr class="separator">

<h4 class="tutorial-segment-title">Quick Links and Resources</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>IronXL Documentation Reference</h3>
      <p>Review the API Reference for IronXL to explore more about functions, features, classes, and namespaces available for handling Excel files.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> IronXL Documentation Reference <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100%; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg">
      </div>
    </div>
  </div>
</div>