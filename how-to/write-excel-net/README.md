# Create Excel Functions with .NET

Creating and updating Excel Spreadsheets programmatically is a common requirement in many C# applications. Leveraging the IronXL library simplifies these tasks dramatically. With IronXL, you can efficiently work with Excel files in numerous formats without the need for extensive coding—simply manipulate the cells directly with the values you wish to use.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Excel .NET Manipulation Instructions</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-the-library">Download the Library for Excel .NET</a></li>
        <li><a href="#anchor-3-write-value-in-specific-cell">Assign values to specific cells</a></li>
        <li><a href="#anchor-4-write-static-values-in-a-range">Insert static values across multiple cells</a></li>
        <li><a href="#anchor-5-write-dynamic-values-in-a-range">Insert dynamic values across a cell range</a></li>
        <li><a href="#anchor-6-replace-excel-cell-value">Modify existing values in cells, rows, columns, or ranges</a></li>
      </ul>
    </div>
     <div class="col-sm-6">
      <div class="download-card">
        <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/faq/excel/how-to-work.svg" class="img-responsive learn-how-to-img replaceable-img">
      </div>
    </div>
  </div>
</div>

## Access and Manipulate Excel Files

First, load the Excel file within your project and access the desired Worksheet as shown in the example below:

```cs
// Initialize and load an existing Excel file
WorkBook workBook = WorkBook.Load("path-to-your-excel-file");
```

After loading the file, select the specific Worksheet:

```cs
// Access a specific Worksheet by name
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");
```

This setup prepares the `workSheet` object for data manipulation. You can learn more about loading different file types and accessing Worksheets [here](https://ironsoftware.com/csharp/excel/how-to/load-spreadsheet/).

**Note:** Ensure `IronXL` is included in your project references and the namespace is imported via `using IronXL`.

<hr class="separator">

## Writing to Specific Cells

Employing IronXL, you can modify a specific cell with ease:

```cs
// Assign a new value to a specific cell
workSheet["Cell Identifier"].Value = "Your Assigned Value";
```

Below is an example demonstrating this operation:

```cs
using IronXL;

// Instantiating the WorkBook
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Accessing the Worksheet
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Updating the value of cell A1
workSheet["A1"].Value = "Updated Content";

// Ensuring changes are saved
workBook.SaveAs("updated-sample.xlsx");
```

This code snippet modifies the `A1` cell of `Sheet1` in the `sample.xlsx` file to `Updated Content`.

**Note:** Always remember to save your file to keep the changes.

### Direct Assignment of String Values

To precisely assign string data without conversion:

```cs
// Assigning a direct textual value to a cell
workSheet["A1"].StringValue = "Exact Text";
```

<hr class="separator">

## Inserting Values Across a Range

Assign values efficiently across a range of cells:

```cs
// Define a range and assign a uniform value
workSheet["Start Cell:End Cell"].Value = "Uniform Value";
```

This method fills every cell within the specified range with `Uniform Value`. Here's how this can be executed:

```cs
using IronXL;

// Preparing the workbook and worksheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Assigning a value over a range of rows
workSheet["B2:B9"].Value = "Spread Value";

// Assigning a value over a range of columns
workSheet["C3:C7"].Value = "Spread Value";

// Save your workbook to apply changes
workBook.SaveAs("updated-sample.xlsx");
```

This example demonstrates filling specified ranges in a Worksheet with a consistent value.

<hr class="separator">

## Dynamically Writing to a Range

Dynamic data insertion into various cells can be executed as follows:

```cs
using IronXL;

// Load and set up the Excel file and sheet
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Loop to insert dynamic data across specified cells
for (int i = 2; i <= 7; i++)
{
    workSheet["B" + i].Value = "Dynamic Value " + i;
    workSheet["D" + i].Value = "Dynamic Value " + i;
}

// Commit the changes to the file
workBook.SaveAs("dynamic-sample.xlsx");
```

This example inserts unique dynamic values into specific rows and columns within the Excel file.

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/write-excel-net/1excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/write-excel-net/1excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## Cell Value Replacement

IronXL allows replacing existing values easily within the entire Worksheet or specific sections:

```cs
// Replace old values throughout the Worksheet
workSheet.Replace("Old Value", "New Value");
```

Further detailed replacements are provided through methods targeting specific rows, columns, or cell ranges:

```cs
// Replace value in a specific row
workSheet.Rows[rowIndex].Replace("Old Value", "New Value");

// Replace value in a specific column
workSheet.Columns[columnIndex].Replace("Old Value", "New Value");

// Replace values within a specified cell range
workSheet["Start Cell:End Cell"].Replace("Old Value", "New Value");
```

For a comprehensive example:

```cs
using IronXL;

// Initialization and setup
WorkBook workBook = WorkBook.Load("sample.xlsx");
WorkSheet workSheet = workBook.GetWorkSheet("Sheet1");

// Replacement operations for various worksheet sections
workSheet.Replace("old", "new");
workSheet.Rows[5].Replace("old", "new");
workSheet.Columns[4].Replace("old", "new");
workSheet["A5:H5"].Replace("old", "new");

// Completing changes by saving
workBook.SaveAs("fully-revised-sample.xlsx");
```

Explore further details on Excel .NET application creation in our comprehensive [tutorial](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/).

<hr class="separator">

<h4 class="tutorial-segment-title">Quick Access to Tutorial Resources</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Explore the API Documentation</h3>
      <p>Dive into the IronXL documentation for a complete breakdown of functions, features, namespaces, classes, and enums available.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Read API Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>