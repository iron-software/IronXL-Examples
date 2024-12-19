# C# Parse Excel File

***Based on <https://ironsoftware.com/how-to/c-sharp-parse-excel-file/>***


In C# applications that utilize Excel spreadsheets, it's common to extract and transform spreadsheet data into various formats for analysis. Leveraging IronXL within the C# environment simplifies these tasks, allowing developers to efficiently parse Excel files as illustrated in the steps below.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>C# Open Excel Worksheets</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-ironxl-for-visual-studio">Download IronXL for Visual Studio</a></li>
        <li><a href="#anchor-5-parse-excel-data-into-numeric-values">C# parse Excel data into numeric values</a></li>
        <li><a href="#anchor-6-parse-excel-data-into-boolean-values">Parse data into boolean values</a></li>
        <li><a href="#anchor-7-parse-excel-file-into-c-collections">Parse files into C# Collections including array, datatables, and datasets</a></li>
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

<h2>How to Parse Excel File in C#</h2>

1. Install the IronXL library for handling Excel files.
2. Load your Excel file by using the `Workbook` class.
3. Select the default `Worksheet` from the workbook.
4. Retrieve values from the Excel `Workbook`.
5. Accurately process and display the retrieved values.

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Download IronXL for Visual Studio  

Begin by [Installing IronXL for Visual Studio](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.parse.excel.file.zip), available as a free tool for developers, or you can also install it via [NuGet](https://www.nuget.org/packages/IronXL.Excel).

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">How to Tutorial</h4>

## 2. Load Excel File

Open your C# project and integrate the Excel file using the `WorkBook.Load()` method from IronXL. Specify the Excel file path as shown:

```cs
// Load the Excel file
WorkBook wb = WorkBook.Load("Path");
```

Once loaded, `wb` now contains the workbook instance. You can then proceed to access a specific worksheet.

<hr class="separator">

## 3. Open the Excel Worksheet

Use the `WorkBook.GetWorkSheet()` method to access a specific worksheet by name:

```cs
// Access a specific worksheet
WorkSheet ws = wb.GetWorkSheet("SheetName");
```

Here, `wb` represents the loaded workbook.

<hr class="separator">

## 4. Retrieve Data from the Excel File

Now, you can fetch and parse data from your chosen worksheet. Below is an example showing how to obtain a particular cell value as a string:

```cs
// Retrieve data from a specific cell
string val = ws["Cell Address"].ToString();
```

In the code above, `ws` refers to the worksheet accessed previously. More examples on [reading Excel file data are available here.](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/#sample-read-xls-or-xlsx-file)

<hr class="separator">

## 5. Parse Excel Data Into Numeric Values

Proceed with parsing Excel data. Here's how IronXL helps parse different numeric types from excel data:

<div class="content-table parse-excel-file">
  <table>
    <tbody>
      <tr class="tr-head">
          <th class="tcol1">DataType</th>
          <th class="tcol2">Method</th>
          <th class="tcol3">Explanation</th>
      </tr>
      <tr>
          <td>int</td>
          <td>WorkSheet ["CellAddress"].IntValue</td>
          <td>Parsing Excel cell value into `Int`.</td>
      </tr>
      <tr>
          <td>Int32</td>
          <td>WorkSheet ["CellAddress"].Int32Value</td>
          <td>For parsing Excel cell value into `Int32`.</td>
      </tr>
      <tr>
          <td>Int64</td>
          <td>WorkSheet ["CellAddress"].Int64Value</td>
          <td>When dealing with large numeric values in your projects.</td>
      </tr>
      <tr>
          <td>float</td>
          <td>WorkSheet ["CellAddress"].FloatValue</td>
          <td>Useful for values that require precision beyond the decimal point.</td>
      </tr>
      <tr>
          <td>Double</td>
          <td>WorkSheet ["CellAddress"].DoubleValue</td>
          <td>When needing to retrieve numeric data with enhanced precision.</td>
      </tr>
      <tr>
          <td>Decimal</td>
          <td>WorkSheet ["CellAddress"].DecimalValue</td>
          <td>When precision with extensive decimal places is needed.</td>
      </tr>
    </tbody>
  </table>
</div>

Below, observe an example that utilizes these methods to transform Excel data into numeric values within a C# context.

```cs
/**
Parsing Numeric Values in Excel
anchor-parse-excel-data-into-numeric-values
**/
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Parse the Excel cell value into a string
    string str_val = ws["B3"].Value.ToString();
    // Parse the Excel cell value into an Int32
    Int32 int32_val = ws["G3"].Int32Value;
    // Parse the Excel cell value into Decimal
    decimal decimal_val = ws["E5"].DecimalValue;

    Console.WriteLine("String from B3: {0}", str_val);
    Console.WriteLine("Int32 from G3: {0}", int32_val);
    Console.WriteLine("Decimal from E5: {0}", decimal_val);
    Console.ReadKey();
}
```

These functions provide outputs as displayed in the image links below:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/1output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/1output.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

And, observe the data from the Excel file `sample.xlsx` here:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/1excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/1excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">
Proceed to the next sections for further parsing methods such as Boolean values and C# collections, including parsing into arrays, datatables, and datasets. 

## 6. Parse Excel Data into Boolean Values

To convert Excel values into boolean data with IronXL, you have access to the `BoolValue` method:

```cs
/**
Conversion to Boolean Values
anchor-parse-excel-data-into-boolean-values
**/
bool val = ws["Cell Address"].BoolValue;
```

In this example, `ws` refers to the previously accessed worksheet. This function will return the boolean representation (e.g., `True` or `False`). Ensure that values in your Excel are in `0`, `1`, `True`, or `False` formats to properly convert to Boolean.

<b>Note: For accurate boolean data type parsing, make sure your Excel file values correspond correctly.</b>

<hr class="separator">

## 7. Parse Excel File into C# Collections

IronXL also affords the transformation of Excel data into several types of C# collections:

<div class="content-table parse-excel-file">
  <table>
    <tbody>
      <tr class="tr-head">
          <th class="tcol1">DataType</th>
          <th class="tcol2">Method</th>
          <th class="tcol3">Explanation</th>
      </tr>
      <tr>
          <td>Array</td>
          <td>WorkSheet ["From:To"].ToArray()</td>
          <td>This method converts specified cell range data into an array format.</td>
      </tr>
      <tr">
          <td>DataTable</td>
          <td>WorkSheet.ToDataTable()</td>
          <td>Converts an entire Excel worksheet into a DataTable, facilitating data manipulation.</td>
      </tr>
      <tr">
          <td>DataSet</td>
          <td>WorkBook.ToDataSet()</td>
          <td>Transforms an entire Excel workbook into a DataSet, where each worksheet is rendered as a DataTable.</td>
      </tr>
    </tbody>
  </table>
</div>

Explore how each collection type is parsed from Excel data:

### 7.1. Parse Excel Data Into Array

For parsing specified ranges into an array, use the following method:

```cs
var array = ws["From:To"].ToArray();
```

To access items specifically, index into the array:

```cs
string item = array[ItemIndex].Value.ToString();
```

Here's an example of array parsing and item selection:

```cs
/**
Array Parsing
anchor-parse-excel-data-into-array
**/
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    var array = ws["B6:F6"].ToArray();
    int itemCount = array.Count();
    string totalItems = array[0].Value.ToString();
    Console.WriteLine("First item in the array: {0}", itemCount);
    Console.WriteLine("Total items from B6 to F6: {0}", totalItems);
    Console.ReadKey();
}
```

The output will be formatted as shown in the image below:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2output.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

The data range from the Excel file `sample.xlsx` is represented in the images:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

### 7.2. Parse Excel Worksheet into DataTable

Converting an Excel worksheet into a DataTable allows for comprehensive data interaction. Here’s how to utilize IronXL for this:

```cs
DataTable dt = ws.ToDataTable();
```

For configuring DataTable column names based on the first row of the Excel file, adjust the parameter as follows:

```cs
DataTable dt = ws.ToDataTable(true);
```

This parameter determines whether the first row in Excel acts as the column names. Detailed information on leveraging ExcelWorksheet as DataTable in C# can be found at [this resource](https://ironsoftware.com/csharp/excel/#excel-sql-datatable).

Consider this example for parsing into DataTable:

```cs
/**
DataTable Parsing
anchor-parse-excel-worksheet-into-datatable
**/
using IronXL;
using System.Data;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Convert Sheet1 of the Excel file into a DataTable with column names
    DataTable dt = ws.ToDataTable(true);
}
```

### 7.3. Parse Excel File into DataSet

For parsing an entire Excel file into a DataSet, wherein each worksheet is represented as a DataTable, use the following:

```cs
DataSet ds = WorkBook.ToDataSet();
```

<b>Note: In this scenario, each of the Workbook's worksheets is transformed into a DataTable within the DataSet. </b>

Here’s a quick example of parsing into a DataSet:

```cs
/**
DataSet Parsing
anchor-parse-excel-file-into-dataset
**/
using IronXL;
using System.Data;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    // Convert the workbook into a DataSet
    DataSet ds = wb.ToDataSet();
    // Access the first DataTable, representing the first worksheet
    DataTable dt = ds.Tables[0];
}
```

Further details on working with Excel SQL Datasets are provided at [this resource](https://ironsoftware.com/csharp/excel/#excel-sql-dataset).

<hr class="separator">
<h4 class="tutorial-segment-title">Tutorial Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
    <div class="col-sm-8">
      <h3>Documentation for Excel in C#</h3>
      <p>Utilize the comprehensive IronXL documentation for leveraging Excel functionalities in your C# projects.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Explore Excel in C# Documentation <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>