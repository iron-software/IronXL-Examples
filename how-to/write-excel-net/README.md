# .NET Excel Functions with IronXL

***Based on <https://ironsoftware.com/how-to/write-excel-net/>***


Developing C# applications often includes tasks like creating or updating Excel spreadsheets programmatically. While Excel .NET integration can seem daunting, the IronXL library simplifies these tasks significantly. It allows developers to work seamlessly with Excel files of any format by directly accessing and modifying cells without extensive code.

### Getting Started with IronXL

--------------------------------------

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Instructions for Excel .NET with IronXL</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-the-library">Download the IronXL Library</a></li>
        <li><a href="#anchor-3-write-value-in-specific-cell">Insert values into specific cells</a></li>
        <li><a href="#anchor-4-write-static-values-in-a-range">Enter static data into multiple cells</a></li>
        <li><a href="#anchor-5-write-dynamic-values-in-a-range">Insert dynamic data across a range of cells</a></li>
        <li><a href="#anchor-6-replace-excel-cell-value">Modify existing cell values in rows, columns, or ranges</a></li>
      </ul>
    </div>
     <div class="col-sm-6">
      <div class="download-card">
        <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/faq/excel/how-to-work.svg" class="img-responsive learn-how-to-img replaceable-img">
      </div>
    </div>
  </div>
</div>

## Accessing Excel Files

First, let's open an Excel file within our project and access a specific worksheet using the follwing code snippets:

```cs
// Load the Excel file into the project
WorkBook workBook = WorkBook.Load("path");
```

This code will load the designated Excel file. Next, let's open a Worksheet:

```cs
// Access a specific WorkSheet
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");
```

You can start manipulating the data in the Excel file using `workSheet`. For additional insights on how to load and interact with different spreadsheet formats, visit [this detailed guide](https://ironsoftware.com/csharp/excel/how-to/load-spreadsheet/).

<b>Note: Remember to include the `IronXL` library in your project references and import it using `using IronXL`.</b>

<hr class="separator">

## Modifying a Specific Cell

A fundamental task is modifying the contents of a single cell. Here's how you can do it using IronXL:

```cs
workSheet["Cell Address"].Value = "Assigned Value";
```

To demonstrate, here's the code that updates a specific cell in our C# project:

```cs
using IronXL;

// Load the Excel file
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Access the Worksheet
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Update cell A1
workSheet["A1"].Value = "new value";

// Save the updated file
workBook.SaveAs("sample.xlsx");
```

This sample modifies the `A1` cell in the `Sheet1` worksheet, then saves the changes to the file.

<b>Note: Always save your Excel file after making modifications, as demonstrated above.</b>

### Forcing Exact Value Assignments

When assigning values, IronXL tries to convert them to appropriate data types. To avoid this and ensure values are assigned exactly as specified, use `StringValue`:

```cs
// Force assign the exact string value
workSheet["A1"].StringValue = "4402-12";
```
<hr class="separator">

## Writing to Multiple Cells

To insert values into a range of cells, you can specify the start and end addresses:

```cs
workSheet["From Cell Address:To Cell Address"].Value = "Assigned Value";
```

Below is an example demonstrating how to apply this on multiple cells:

```cs
using IronXL;

// Load the Excel file
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Access the desired Worksheet
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Update a row range
workSheet["B2:B9"].Value = "new value";

// Update a column range
workSheet["C3:C7"].Value = "new value";

// Save the modified Excel file
workBook.SaveAs("sample.xlsx");
```

This code snippet updates specific row and column ranges with new static values.

<hr class="separator">

## Incorporating Dynamic Values

Dynamic data can also be introduced into a range of cells as illustrated in the following code snippet:

```cs
using IronXL;

// Load the file
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Access the Worksheet
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Dynamically update cells across column B and D from row 2 to 7
for (int i = 2; i <= 7; i++) {
    workSheet["B" + i].Value = "Value" + i;
    workSheet["D" + i].Value = "Value" + i;
}

// Save changes
workBook.SaveAs("sample.xlsx");
```

This approach dynamically assigns values in columns `B` and `D` ranging from row 2 to row 7.

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/write-excel-net/1excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/write-excel-net/1excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## Replacing Cell Values

With IronXL, updating existing cell values is straightforward using the `Replace()` method:

```cs
workSheet.Replace("old value", "new value");
```

The code above will substitute `new value` for all occurrences of `old value` throughout the worksheet.

### Specific Row Updates

If the update is confined to a particular row, the approach below is suitable:

```cs
workSheet.Rows[RowIndex].Replace("old value", "new value");
```

### Specific Column Updates

Similarly, updating a specific column is achieved with the following:

```cs
workSheet.Columns[ColumnIndex].Replace("old value", "new value");
```

### Range-Specific Updates

To target a specific range for updates, use:

```cs
workSheet["From Cell Address : To Cell Address"].Replace("old value", "new value");
```

Here's how you might utilize all these methods to update values efficiently:

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWork Sheet("Sheet1");

// Update the whole WorkSheet
workSheet.Replace("old", "new");

// Target updates in specific row and column
workSheet.Rows[5].Replace("old", "new");
workSheet.Columns[4].Replace("old", "new");

// Apply changes to a specific range
workSheet["A5:H5"].Replace("old", "new");

// Save all changes
workBook.SaveAs("sample.xlsx");
```

For comprehensive guidance on creating .NET applications that operate with Excel files, explore our full tutorial on [how to open and write to Excel files in C#](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/).

<hr class="separator">

### Quick Access to IronXL Documentation

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Access Our Complete API Reference</h3>
      <p>Delve into the IronXL documentation for a comprehensive list of functions, features, namespaces, classes, and enums available for your projects.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> Explore API Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>