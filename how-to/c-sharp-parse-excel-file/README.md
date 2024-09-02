# C# Parse Excel File 

Parsing Excel file data in C# can often be requisite when building applications that require data analysis and transformation. IronXL simplifies this task immensely, leveraging its robust features in the C Sharp environment. Follow the outlined steps below for a comprehensive guide on parsing Excel data into various suitable formats. 

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>C# Open Excel Worksheets</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-ironxl-for-visual-studio">Download IronXL for Visual Studio</a></li>
        <li><a href="#anchor-5-parse-excel-data-into-numeric-values">Parse Excel data into numeric values</a></li>
        <li><a href="#anchor-6-parse-excel-data-into-boolean-values">Transform data into boolean values</a></li>
        <li><a href="#anchor-7-parse-excel-file-into-c-collections">Convert files into C# Collections including arrays, datatables, and datasets</a></li>
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

1. Install the IronXL library to facilitate Excel file processing.
2. Open the Excel file by creating a `Workbook` instance.
3. Select the default `Worksheet`.
4. Extract values from the `Workbook`.
5. Process and display these values accurately.

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Download IronXL for Visual Studio  

Begin by [downloading IronXL for Visual Studio](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.parse.excel.file.zip) or install it via [NuGet](https://www.nuget.org/packages/IronXL.Excel), both are viable options for your development needs. 

```shell
Install-Package IronXL.Excel
```
<hr class="separator">
<h4 class="tutorial-segment-title">How to Tutorial</h4>

## 2. Load Excel File

Start by loading your Excel file into your project by using `WorkBook.Load()` from IronXL, specify the file path like so:

```cs
//Load Excel file
WorkBook wb = WorkBook.Load("Path");
```

`wb` will hold your loaded Excel workbook, allowing us to proceed to work with worksheets.

<hr class="separator">

## 3. Open the Excel Worksheet

To open a specific worksheet, enact the `WorkBook.GetWorkSheet()` method as follows:

```cs
//Specify Worksheet
WorkSheet ws = Wb.GetWorkSheet("SheetName");
```

`wb` refers to the Workbook instance created earlier.

<hr class="separator">

## 4. Retrieve Data from Excel File

Retrieving data from an Excel worksheet is straightforward in IronXL. Here’s how to access a specific cell and convert its value to a string:

```cs
//Access Data by Cell
string val = ws["Cell Address"].ToString();
```

`ws` represents the Worksheet, allowing easy access to its cells.

<hr class="separator">

## 5. Parse Excel Data Into Numeric Values

Next, let’s interpret numeric Excel data into programmatically usable formats with IronXL’s parsing methods:

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
          <td>Converts cell value to `Int` for integer operations.</td>
      </tr>
      <tr>
          <td>Int32</td>
          <td>WorkSheet ["CellAddress"].Int32Value</td>
          <td>Converts cell value to `Int32` useful for larger integer values.</td>
      </tr>
      <tr>
          <td>Int64</td>
          <td>WorkSheet ["CellAddress"].Int64Value</td>
          <td>Useful for extremely large integer values.</td>
      </tr>
      <tr>
          <td>float</td>
          <td>WorkSheet ["CellAddress"].FloatValue</td>
          <td>Ideal for values requiring decimals.</td>
      </tr>
      <tr>
          <td>Double</td>
          <td>WorkSheet ["CellAddress"].DoubleValue</td>
          <td>Best for highly precise numeric data.</td>
      </tr>
      <tr>
          <td>Decimal</td>
          <td>WorkSheet ["CellAddress"].DecimalValue</td>
          <td>Optimal for extensive precision and scale, especially in financial data.</td>
      </tr>
    </tbody>
  </table>
</div>

Here’s an example using these functions to parse and interact with Excel data:

```cs
/**
 * Parse into Numeric Values
 * anchor-parse-excel-data-into-numeric-values
 **/
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    //Parse Excel cell value into string
    string str_val = ws ["B3"].Value.ToString();
    //Parse Excel cell value into Int32
    Int32 int32_val = ws ["G3"].Int32Value;
    //Parse Excel cell value into Decimal
    decimal decimal_val = ws ["E5"].DecimalValue;
 
    Console.WriteLine("Parse B3 Cell Value into String: {0}", str_val);
    Console.WriteLine("Parse G3 Cell Value into Int32: {0}", int32_val);
    Console.WriteLine("Parse E5 Cell Value into Decimal: {0}", decimal_val);
    Console.ReadKey();
}
```

This code will display the parsed values from the file `sample.xlsx` as shown:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/1output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/1output.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

And here are the visualized values from the Excel file `sample.xlsx`:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/1excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/1excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## 6. Convert Excel Data into Boolean Values

For conversion of Excel data into Boolean values, IronXL provides the `BoolValue` function as shown here:

```cs
/**
 * Convert into Boolean Values
 * anchor-parse-excel-data-into-boolean-values
 **/
bool Val = ws ["Cell Address"].BoolValue;
```

`ws` represents the previously identified Worksheet. This function will interpret the cell's value as either `True` or `False`.

<b>Note: To convert cell values into booleans, ensure the Excel values are in (`0` , `1`) or (`True` , `False`) formats.</b>

<hr class="separator">

## 7. Convert Excel File Data into C# Collections

IronXL enables conversion of Excel data into various C# collection types, facilitating easy data handling and manipulation:

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
          <td>A convenient function for data range conversion to array format.</td>
      </tr>
      <tr>
          <td>DataTable</td>
          <td>WorkSheet.ToDataTable()</td>
          <td>Converts the full Excel worksheet into a DataTable, facilitating structured data use.</td>
      </tr>
      <tr>
          <td>DataSet</td>
          <td>WorkBook.ToDataSet()</td>
          <td>Translates an entire Excel workbook into a DataSet, converting sheets into DataTables.</td>
      </tr>
    </tbody>
  </table>
</div>

Here we will observe each method to convert Excel data accordingly.

### 7.1. Convert Excel Data Into Array

IronXL facilitates data conversion from specified cell ranges into arrays, as demonstrated below:

```cs
/**
 * Convert into Array
 * anchor-parse-excel-data-into-array
 **/
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    var array = ws ["B6:F6"].ToArray();
    int item = array.Length;
    string total_items = array [0].Value.ToString();
    Console.WriteLine("First item in the array: {0}", total_items);
    Console.WriteLine("Total items from B6 to F6: {0}", item);
    Console.ReadKey();
}
```

Output for the above code is provided as follows:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2output.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

The range values of the Excel file `sample.xlsx` from `B6` to `F6` are illustrated here:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-parse-excel-file/2excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

### 7.2. Convert Excel Worksheet Into DataTable

The ease of IronXL extends to converting specific Excel Worksheets directly into DataTables using `.ToDataTable()` function:

```cs
DataTable dt = WorkSheet.ToDataTable();
```

Set the first row of the Excel file as the column names in the DataTable by using:

```cs
DataTable dt = WorkSheet.ToDataTable(True);
```

Explore more on handling [ExcelWorksheet as DataTable in C#](https://ironsoftware.com/csharp/excel/#excel-sql-datatable) for enhanced usability.

And here is how you would parse a Worksheet into a DataTable:

```cs
/**
 * Convert into DataTable
 * anchor-parse-excel-worksheet-into-datatable
 **/
using IronXL;
using System.Data;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Parse Sheet1 into a DataTable, setting the first row of the Excel file as column names
    DataTable dt = ws.ToDataTable(true); 
}
```

### 7.3. Convert an Entire Excel File into a DataSet

To convert an entire Excel file into a DataSet, leverage IronXL’s `.ToDataSet()` method:

```cs
/**
 * Convert File into DataSet
 * anchor-parse-excel-file-into-dataset
 **/
using IronXL;
using System.Data;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    // Convert the WorkBook into a DataSet
    DataSet ds = wb.ToDataSet(); 
    // Retrieve a DataTable, originally a WorkSheet, from the DataSet
    DataTable dt = ds.Tables[0];
}
```

<b>Note: Converting an Excel file into a DataSet will convert all WorkSheets into DataTables within the DataSet.</b>

You can learn more on integrating and manipulating [Excel as SQL Dataset](https://ironsoftware.com/csharp/excel/#excel-sql-dataset) in your applications.

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
      <p>Explore the comprehensive IronXL documentation for handling Excel files in C#. It includes detailed guidance on functionality, classes, namespaces, and more for your projects.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> Explore Excel Documentation in C# <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>