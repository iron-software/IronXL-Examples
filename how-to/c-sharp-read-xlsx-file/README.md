# C# Read XLSX File

***Based on <https://ironsoftware.com/how-to/c-sharp-read-xlsx-file/>***


Handling multiple Excel document formats often necessitates the reading and manipulation of data through programming. This guide demonstrates how to read data from an Excel spreadsheet using C# with the help of the IronXL library.

```html
<div class="learn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Read .XLSX Files in C#</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-get-ironxl-for-your-project">Add IronXL to your project</a></li>
        <li><a href="#anchor-2-load-workbook">Open a WorkBook</a></li>
        <li><a href="#anchor-4-access-data-from-worksheet">Retrieve data from a WorkSheet</a></li>
        <li><a href="#anchor-5-perform-functions-on-data">Implement functions like Sum, Min, and Max</a></li>
        <li><a href="#anchor-6-read-excel-worksheet-as-datatable">Convert a WorkSheet to a DataTable, DataSet, and beyond</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/faq/excel/how-to-work.svg" class="img-responsive learn-how-to-img replaceable-img">
      </div>
    </div>
  </div>
</div>
```

<hr class="separator">

### Step 1

## Get IronXL for Your Project

Integrate IronXL into your application to easily handle Excel file formats in C#. Install IronXL by [downloading it directly](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.read.xlsx.zip) or through [NuGet in Visual Studio](https://www.nuget.org/packages/IronXL.Excel). This tool is free to use during development.

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
### Tutorial Instructions

## Load a WorkBook

`WorkBook` is the core class in IronXL which allows full manipulation of the Excel workbook. To open an Excel document, you can use:

```csharp
// Example Code to Load an Excel workbook
WorkBook wb = WorkBook.Load("sample.xlsx");
```

Here, the `WorkBook.Load()` method is used to open `sample.xlsx` and assign it to `wb`, which can then be utilized to perform various operations on the workbook.

<hr class="separator">

## Access a Specific WorkSheet

IronXL's `WorkSheet` class enables access to individual sheets. It provides multiple ways to retrieve a sheet:

```csharp
// Access the worksheet by sheet name
WorkSheet ws = wb.GetWorkSheet("Sheet1");
```

```csharp
// Access the worksheet by index
WorkSheet ws = wb.WorkSheets[0];
```

```csharp
// Access the default worksheet
WorkSheet ws = wb.DefaultWorkSheet;
```

```csharp
// Access the first available worksheet
WorkSheet ws = wb.WorkSheets.First();
```

```csharp
// Access the first or default worksheet
WorkSheet ws = wb.WorkSheets.FirstOrDefault();
```

Once `ws` (a `WorkSheet` instance) is obtained, it offers functionalities to access and manipulate each cell's data.

<hr class="separator">

## Retrieve Data from a WorkSheet

To extract data from a worksheet, IronXL provides straightforward methods to access cell values:

```csharp
string stringValue = ws["cell address"].ToString(); // Retrieving string data
int integerValue = ws["cell address"].Int32Value; // Retrieving integer data

// Iterate over a range of cells
foreach (var cell in ws["A2:A10"])
{
    Console.WriteLine($"Value is: {cell.Text}");
}
```

Here is a complete example that demonstrates accessing data from a worksheet:

```csharp
// Code Snippet: Accessing WorkSheet Data
using IronXL;

static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");

    foreach (var cell in ws["A2:A10"])
    {
        Console.WriteLine($"Value is: {cell.Text}");
    }

    Console.ReadKey();
}
```

Below are some snapshots of the process in action using `Sample.xlsx`:

<center>
  <div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-input1.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-input1.png" alt="" class="img-responsive add-shadow"></a>
  </div>
</center>

<hr class="separator">

## Apply Functions on Data

Extracting specific data with functions like Sum, Min, and Max is quite straightforward:

```csharp
// Aggregate functions implementation
decimal sumValue = ws["From:To"].Sum();
decimal minValue = ws["From:To"].Min();
decimal maxValue = ws["From:To"].Max();
```

For more detailed techniques on performing operations on Excel data, reference our comprehensive tutorial on [Writing and Reading Excel Files with C#](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/#advanced-operations-sum-avg-count-etc).

```csharp
// Code Snippet: Sum, Min, Max Functions
using IronXL;

static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");

    decimal sumResult = ws["G2:G10"].Sum();
    decimal minResult = ws["G2:G10"].Min();
    decimal maxResult = ws["G2:G10"].Max();

    Console.WriteLine($"Sum is: {sumResult}");
    Console.WriteLine($"Min is: {minResult}");
    Console.WriteLine($"Max is: {maxResult}");

    Console.ReadKey();
}
``` 

This output and Excel file visualizes the processed results:

<center>
  <div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-output2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-output2.png" alt="" class="img-responsive add-shadow"></a>
  </div>
</center>

And the original `Sample.xlsx` file looks like this:

<center>
  <div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-2.png" target="_blank"><img[src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-2.png" alt="" class="img-responsive add-shadow"></a>
  </div>
</center>

<hr class="separator">

## Convert Excel WorkSheet to DataTable

Converting a WorkSheet to a DataTable with IronXL is a straightforward task:

```csharp
// Convert a worksheet to a DataTable
DataTable dataTable = WorkSheet.ToDataTable();
```

If the first row should serve as column names:

```csharp
// Include the first row as column names in the DataTable conversion
DataTable dataTable = WorkSheet.ToDataTable(True);
```

The Boolean parameter in `ToDataTable()` determines whether the first row is used as column names, defaulting to `False`.

```csharp
// Example: WorkSheet as DataTable
using IronXL;
using System.Data;

static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    DataTable dt = ws.ToDataTable(true);

    foreach (DataRow row in dt.Rows)
    {
        for (int i = 0; i < dt.Columns.Count; i++)
        {
            Console.Write(row[i]);
        }
    }
}
```

Using the code above, every cell value in the WorkSheet can be accessed and utilized as needed.

<hr class="separator">

## Convert Excel WorkBook to DataSet

IronXL also simplifies the process of using a complete Excel WorkBook as a DataSet:

```csharp
// Convert a WorkBook to a DataSet
DataSet dataSet = WorkBook.ToDataSet();
```

In the following example, we see how to use an Excel file as a DataSet:

```csharp
// Example: Excel File as DataSet
using IronXL;
using System.Data;

static void Main(string[] args)
{           
    WorkBook wb = WorkBook.Load("sample.xlsx");
    DataSet ds = wb.ToDataSet(); // Parse the WorkBook into a DataSet

    foreach (DataTable dt in ds.Tables)
    {
        Console.WriteLine(dt.TableName);
    }
}
```

The output and visualization of the Excel file `Sample.xlsx`:

<center>
  <div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc10-output2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc10-output2.png" alt="" class="img-responsive add-shadow"></a>
  </div>
</center>

<center>
  <div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc10-2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc10-2.png" alt="" class="img-responsive add-shadow"></a>
  </div>
</center>

The example demonstrates easy parsing of an Excel file into a DataSet, allowing manipulation of each WorkSheet as a DataTable. For more insights and code examples on parsing Excel as a DataSet, explore our [dedicated guide](https://ironsoftware.com/csharp/excel/#excel-sql-dataset).

Let's examine another snippet showing how to access each cell value across all ExcelSheets:

```csharp
// Example: WorkSheet Cell Values
using IronXL;
using System.Data;

static void Main(string[] args)
{ 
WorkBook wb = WorkBook.Load("sample.xlsx");
DataSet ds = wb.ToDataSet(); // Treat the entire Excel file as a DataSet

foreach (DataTable dt in ds.Tables) // Treat each Excel WorkSheet as a DataTable
{
    foreach (DataRow row in dt.Rows) // Iterate through each sheet's rows
    {
        for (int i = 0; i < dt.Columns.Count; i++) // Iterate through each row's columns
        {
            Console.Write(row[i]);
        }
    }
}
}
```

This example effectively illustrates how simple it is to access each cell value across every worksheet within an Excel file.

For further details on how to [Read Excel Files Without Interop](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/), you can find additional code examples here.

<hr class="separator">

### Tutorial Quick Access

```html
<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>API Reference for IronXL</h3>
      <p>Explore detailed information about IronXL's features, classes, method fields, namespaces, and enums in our documentation.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> API Reference for IronXL <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>
```