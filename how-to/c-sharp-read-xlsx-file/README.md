# C# Read XLSX File

***Based on <https://ironsoftware.com/how-to/c-sharp-read-xlsx-file/>***


Dealing with different Excel file formats typically involves reading and manipulating the data with C# programming. In this tutorial, we'll explore how to read information from an Excel spreadsheet using the IronXL library.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Handling .XLSX Files in C#</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-get-ironxl-for-your-project">Integrate IronXL into your project</a></li>
        <li><a href="#anchor-2-load-workbook">Open a workbook</a></li>
        <li><a href="#anchor-4-access-data-from-worksheet">Retrieve data from a worksheet</a></li>
        <li><a href="#anchor-5-perform-functions-on-data">Perform calculations like Sum, Min, & Max</a></li>
        <li><a href="#anchor-6-read-excel-worksheet-as-datatable">Convert a worksheet into a DataTable, DataSet, and more</a></li>
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

## 1. Incorporate IronXL into Your Project

For easy manipulation of Excel files in C#, include IronXL in your project. You can [download IronXL directly](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.read.xlsx.zip) or use [NuGet to install it via Visual Studio](https://www.nuget.org/packages/IronXL.Excel). IronXL is free for development purposes.

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">How to Tutorial</h4>

## 2. Open a Workbook

The `WorkBook` class in IronXL allows full control over the Excel files. For instance, to open a workbook, you would utilize:

```cs
// Load Workbook
WorkBook wb = WorkBook.Load("sample.xlsx");  // Path to your Excel file
```
In the snippet above, the `WorkBook.Load()` method opens `sample.xlsx` and assigns it to `wb`. You can then operate on `wb` using any of its worksheets.

<hr class="separator">

## 3. Access a Specific Worksheet

The `WorkSheet` class in IronXL facilitates worksheet handling in various ways:

```cs
// Access Sheet by Name
WorkSheet ws = wb.GetWorkSheet("Sheet1");  // Access using sheet name
```
`wb` is a previously declared WorkBook instance.

OR
```cs
// Access Sheet by Index
WorkSheet ws = wb.WorkSheets[0];  // Access by sheet index
```
OR

```cs
// Access Default WorkSheet
WorkSheet ws = wb.DefaultWorkSheet;  // Access the default sheet
```
OR

```cs
// Access First WorkSheet
WorkSheet ws = wb.WorkSheets.First();  // Access the first sheet
```
OR

```cs
// Access First or Default WorkSheet
WorkSheet ws = wb.WorkSheets.FirstOrDefault();  // Access the first or default sheet
```
Once you access a worksheet `ws`, you can retrieve data from it and perform any Excel operations you need.

<hr class="separator">

## 4. Retrieve Data from a Worksheet

Retrieving data from a `WorkSheet` can be accomplished as follows:

```cs
string cellData = ws["cell address"].ToString();  // Retrieve string data
int intValue = ws["cell address"].Int32Value;  // Retrieve integer data
```

It's also possible to extract values from multiple cells within a specific column:

```cs
foreach (var cell in ws["A2:A10"])
{
    Console.WriteLine("value is: {0}", cell.Text);
}
```
This will print the values from cells `A2` to `A10`.

Here's a complete example demonstrating the steps above:

```cs
// Access WorkSheet Data
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");  // Load the Excel file
    WorkSheet ws = wb.GetWorkSheet("Sheet1");  // Access the specific worksheet
    foreach (var cell in ws["A2:A10"])  // Iterate through cells from A2 to A10
    {
        Console.WriteLine("value is: {0}", cell.Text);  // Print out each cell's text
    }
    Console.ReadKey();
}
```

The output will appear as follows:

<center>
<div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-input1.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-input1.png" alt="" class="img-responsive add-shadow"></a>
</div>
</center>

With the Excel file `Sample.xlsx` showcased:

<center>
<div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-1.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-1.png" alt="" class="img-responsive add-shadow"></a>
</div>
</center>

This method clearly demonstrates the simplicity and power of managing Excel data within your projects using these techniques.

<hr class="separator">

## 5. Apply Functions to Data

Accessing and applying aggregate functions like Sum, Min, or Max on data from an Excel `WorkSheet` is straightforward:

```cs
decimal sumValue = ws["From:To"].Sum();  // Calculate sum
decimal minValue = ws["From:To"].Min();  // Find minimum
decimal maxValue = ws["From:To"].Max();  // Find maximum
```

For more advanced usage, refer to our comprehensive guide on [Writing C# Excel Files](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/#advanced-operations-sum-avg-count-etc) which includes detailed examples on aggregate functions.

```cs
// Sum, Min, Max Functions
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");  // Load our sample Excel file
    WorkSheet ws = wb.GetWorkSheet("Sheet1");  // Access the first worksheet

    decimal sum = ws["G2:G10"].Sum();  // Calculate sum for cells G2 through G10
    decimal min = ws["G2:G10"].Min();  // Find minimum in range G2 through G10
    decimal max = ws["G2:G10"].Max();  // Find maximum in range G2 through G10

    Console.WriteLine("Sum is: {0}", sum);  // Display the sum
    Console.WriteLine("Min is: {0}", min);  // Display the minimum
    Console.WriteLine("Max is: {0}", max);  // Display the maximum
    Console.ReadKey();
}
```
This code will produce the following display:

<center>
<div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-output2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-output2.png" alt="" class="img-responsive add-shadow"></a>
</div>
</center>

And this represents the data from the Excel file `Sample.xlsx`:

<center>
<div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-2.png" alt="" class="img-responsive add-shadow"></a>
</div>
</center>

<hr class="separator">

## 6. Convert Excel Worksheet to DataTable

Utilizing IronXL to handle an Excel `WorkSheet` as a `DataTable` is straightforward:

```cs
DataTable dt = WorkSheet.ToDataTable();  // Convert to DataTable
```

To use the first row of the Excel sheet as the column names for the `DataTable`:

```cs
DataTable dt = WorkSheet.ToDataTable(true);  // Use first row as column names
```
By setting the Boolean parameter of the `ToDataTable()` method, you determine if the first row should act as the column names, with the default being `False`.

```cs
// WorkSheet as DataTable
using IronXL;
using System.Data;
static void Main(string[] args)
{          
    WorkBook wb = WorkBook.Load("sample.xlsx");  // Load the workbook
    WorkSheet ws = wb.GetWorkSheet("Sheet1");  // Get the first worksheet
    DataTable dt = ws.ToDataTable(true);  // Convert the worksheet into a DataTable, using the first row as column names

    foreach (DataRow row in dt.Rows)  // Iterate through each row
    {
        for (int i = 0; i < dt.Columns.Count; i++)  // Iterate through each column in the row
        {
            Console.Write(row[i]);  // Print out each cell value
        }

    }
}
```
This code allows direct access to every cell value in the worksheet, using it as needed.

<hr class="separator">

## 7. Convert Excel File to DataSet

IronXL also facilitates the conversion of an entire Excel file (`WorkBook`) into a `DataSet`, allowing further manipulation:

```cs
DataSet ds = WorkBook.ToDataSet();  // Convert the entire workbook to a DataSet
```
Explained through the following example:

```cs
// Excel File as DataSet
using IronXL;
using System.Data;
static void Main(string[] args)
{          
    WorkBook wb = WorkBook.Load("sample.xlsx");  // Load the Excel file into a WorkBook
    DataSet ds = wb.ToDataSet();  // Convert the entire WorkBook to a DataSet

    foreach (DataTable dt in ds.Tables)  // Iterate through each DataTable in the DataSet
    {
        Console.WriteLine(dt.TableName);  // Print the name of each DataTable
    }
}
```
The output will appear like this:

<center>
<div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc10-output2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc10-output2.png" alt="" class="img-responsive add-shadow"></a>
</div>
</center>

And this will show how the Excel file `Sample.xlsx` looks:

<center>
<div class="center-image-wrapper">
    <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc10-2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc10-2.png" alt="" class="img-responsive add-shadow"></a>
</div>
</center>

In this example, we've demonstrated the simple, yet flexible ways to handle an Excel file as a DataSet, utilizing each worksheet as a DataTable and manipulating the data accordingly. For more details on parsing Excel as a DataSet, dive into our resources [here](https://ironsoftware.com/csharp/excel/#excel-sql-dataset) which include broader code examples.

Let's explore another example on how to access cell values in all Excel sheets. You can easily retrieve values from every worksheet in the Excel file:

```cs
// WorkSheet Cell Values
using IronXL;
using System.Data;
static void Main(string[] args)
{ 
WorkBook wb = WorkBook.Load("sample.xlsx");  // Load the entire Excel file as a DataSet
DataSet ds = wb.ToDataSet();  // Treat the complete Excel file as a DataSet
foreach (DataTable dt in ds.Tables)  // Treat each Excel Worksheet as a DataTable
{
    foreach (DataRow row in dt.Rows)  // Iterate through each row in the DataTable
    {
        for (int i = 0; i < dt.Columns.Count; i++)  // Iterate through each column in the row
        {
            Console.Write(row[i]);  // Print out each cell value
        }
    }
}
}
```
This example showcases the convenience of accessing individual cell values from every sheet in the Excel file.

For comprehensive guides on [Reading Excel Files Without Interop](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/) and further examples, check out our resources on this topic.

<hr class="separator">

<h4 class="tutorial-segment-title">Tutorial Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>API Reference for IronXL</h3>
      <p>Explore more about IronXL's features, classes, methods, fields, namespaces, and enums through our documentation.</p>
      <a class="doc-link" href="/csharp/excel/object-reference/api/" target="_blank"> API Reference for IronXL <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>

This comprehensive tutorial provides a deep dive into working with Excel files using C# and IronXL, from basic reading to advanced data manipulation techniques, all demonstrated with clear examples and supportive documentation.