# Importing Excel Files in C#

***Based on <https://ironsoftware.com/how-to/csharp-import-excel/>***


As developers, we frequently encounter the need to extract and utilize data from Excel files to meet the demands of our applications and data processing tasks. Using IronXL, this process is streamlined, enabling us to import data directly into a C# project and manipulate it effortlessly.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Import Excel Data in C#</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-import-data-with-the-ironxl-library">Import Data Using IronXL</a></li>
        <li><a href="#anchor-3-import-excel-data-in-c-num">Work With Excel Data in C#</a></li>
        <li><a href="#anchor-4-import-excel-data-of-specific-range">Handle Data from Specific Cells</a></li>
        <li><a href="#anchor-5-import-excel-data-by-aggregate-functions">Use Excel Aggregate Functions</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/faq/excel/how-to-work.svg" class="img-responsive learn-how-to-img replaceable-img">
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Import Data Using IronXL

Use the conveniently provided functions of the IronXL Excel library to import data. This tool is freely available for development usage.

Install it in your [C# Project](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.Import.Excel.Csharp.zip) or through the provided [NuGet package](https://www.nuget.org/packages/IronXL.Excel).

```shell
Install-Package IronXL.Excel
```


<hr class="separator">
<h4 class="tutorial-segment-title">Gallery Guide</h4>

## 2. Access the Worksheet for Your Project

To begin, load your Excel workbook into your C# application using the IronXL library:

```cs
// Load the Excel file
WorkBook wb = WorkBook.Load("Path");
```

Now, grab a specific worksheet from the loaded workbook:

```cs
// Access a particular worksheet by name
WorkSheet ws = wb.GetWorkSheet("SheetName");
```

Here are alternate ways to access worksheets in your project:

```cs
// Alternative methods to access worksheets
WorkBook.WorkSheets[SheetIndex];   // By sheet index
WorkBook.DefaultWorkSheet;         // Default worksheet
WorkBook.WorkSheets.First();       // First worksheet
WorkBook.WorkSheets.FirstOrDefault();  // First or default worksheet
```

Let's explore various methods to import and interact with Excel data.

<hr class="separator">

## 3. Operating with Excel Data in C&num;

Simply refer to the cells directly to import data using their addresses:

```cs
WorkSheet["Cell Address"];  // Addressing directly
```

Alternatively, retrieve cell data using row and column indexes:

```cs
WorkSheet.Rows[RowIndex].Columns[ColumnIndex];  // By indexes
```

Assign the extracted cell values to variables:

```cs
// Assign cell data to a variable
string value = WorkSheet["Cell Address"].ToString();
string indexedValue = WorkSheet.Rows[RowIndex].Columns[ColumnIndex].Value.ToString();  // Using indexes
```

These methods initiate at the zero index for both rows and columns.

<hr class="separator">

## 4. Extract Data from a Specific Range

Specify the range of data within the Excel workbook that you wish to import:

```cs
WorkSheet["starting Cell Address : Ending Cell Address"];  // Specify range
```

Discover more techniques [here](https://ironsoftware.com/csharp/excel/#excel-ranges) for managing data across different ranges.

```cs
// Example using IronXL to import specific cell data
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    string value = ws["A4"].Value.ToString();  // Import a specific cell
    Console.WriteLine("Imported value from cell A4: {0}", value);
    Console.WriteLine("Values in the range from B3 to B9:");
    // Loop through cells in a range
    foreach (var item in ws["B3:B9"])
    {
        Console.WriteLine(item.Value.ToString());
    }
    Console.ReadKey();
}
```

The example code accurately displays the desired result based on the content from the file `sample.xlsx`.

### Visualization of Excel File and Output 

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-import-excel/1output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-import-excel/1output.png" alt="" class="img-responsive add-shadow"></a>
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-import-excel/1excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-import-excel/1excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## 5. Employing Aggregate Functions

Apply aggregate functions to extract and compute data from Excel files:

* **Sum**
```cs
WorkSheet["Starting Cell Address : Ending Cell Address"].Sum();  // Sum up data in range
```

* **Average**
```cs
WorkSheet["Starting Cell Address : Ending Cell Address"].Avg();  // Calculate average
```

* **Min and Max**
```cs
WorkSheet["Starting Cell Address : Ending Cell Address"].Min();  // Minimum value
WorkSheet["Starting Cell Address : Ending Cell Address"].Max();  // Maximum value
```

For further insights into using aggregate functions in your projects, refer to the [detailed guide](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/#advanced-operations-sum-avg-count-etc).

Here's a practical example:

```cs
// Using aggregate functions to import data
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    decimal Sum = ws["D2:D9"].Sum();
    decimal Avg = ws["D2:D9"].Avg();
    decimal Min = ws["D2:D9"].Min();
    decimal Max = ws["D2:D9"].Max();
    Console.WriteLine("Sum From D2 To D9: {0}", Sum);
    Console.WriteLine("Average From D2 To D9: {0}", Avg);
    Console.WriteLine("Minimum From D2 To D9: {0}", Min);
    Console.WriteLine("Maximum From D2 To D9: {0}", Max);
    Console.ReadKey();
}
```

### Aggregate Function Results Display 

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-import-excel/2output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-import-excel/2output.png" alt="" class="img-responsive add-shadow"></a>
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-import-excel/2excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-import-excel/2excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## 6. Comprehensive Extraction of Excel Data

To incorporate entire datasets from an Excel file into a C# project, first convert the loaded workbook into a `DataSet`, where each worksheet becomes a `DataTable`:

```cs
// Convert an entire workbook into a data set
DataSet ds = WorkBook.ToDataSet();
```

Additionally, setting the first column of an Excel sheet as the `DataTable` column name is conveniently achievable by adjusting the `ToDataSet()` function:

```cs
ToDataSet(true);  // Set the first column as DataTable column name
```

A complete example of importing Excel data into a dataset and setting the first column as data table column name is as follows:

```cs
// Full example demonstrating data import into a DataSet
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    DataSet ds = new DataSet();
    ds = wb.ToDataSet(true);
    Console.WriteLine("Excel file data successfully imported into dataset.");
    Console.ReadKey();
}
```

For a deep dive into working with `Excel Dataset and DataTable`, explore additional [examples and documentation](https://ironsoftware.com/csharp/excel/#read-excel).

<hr class="separator">

<h4 class="tutorial-segment-title">Library Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Discover the IronXL Reference</h3>
      <p>Gain more insights into the methods of managing Excel data using our extensive API documentation.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> Discover the IronXL Reference <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg">
      </div>
    </div>
  </div>
</div>