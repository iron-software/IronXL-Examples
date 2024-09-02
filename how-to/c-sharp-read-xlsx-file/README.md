# C# Read XLSX File

When managing Excel documents, developers frequently need to programmatically read and manipulate data. This tutorial will show you how to extract data from Excel spreadsheets using C# with the aid of the handy IronXL library.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Read .XLSX Files in C#</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-get-ironxl-for-your-project">Set Up IronXL for Your Project</a></li>
        <li><a href="#anchor-2-load-workbook">Initialize a Workbook</a></li>
        <li><a href="#anchor-4-access-data-from-worksheet">Retrieve Data from a Worksheet</a></li>
        <li><a href="#anchor-5-perform-functions-on-data">Implement Functions Like Sum, Min, & Max</a></li>
        <li><a href="#anchor-6-read-excel-worksheet-as-datatable">Convert a Worksheet to a DataTable, DataSet, and More</a></li>
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

## 1. Set Up IronXL for Your Project

Integrate IronXL easily into your C# projects to manage Excel file formats. You may [download IronXL directly](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.read.xlsx.zip) or install it using [NuGet in Visual Studio](https://www.nuget.org/packages/IronXL.Excel). The software is freely available for development purposes.

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">Tutorial Instructions</h4>

## 2. Initialize a Workbook

`WorkBook` is an IronXL class that provides full access to an Excel file and all of its features. For instance, to open an Excel file, you would use:

```cs
// Load a Workbook
WorkBook wb = WorkBook.Load("sample.xlsx"); // Path to Excel file
```

Here, the `WorkBook.Load()` method opens `sample.xlsx` into the variable `wb`. Once loaded, you can execute any operations on `wb` by accessing a specific Worksheet within the Excel file.

<hr class="separator">

## 3. Retrieve Data from a Specific Worksheet

IronXL offers the `WorkSheet` class to facilitate the access to individual sheets within the Excel file:

```cs
// Access Sheet by Name
WorkSheet ws = wb.GetWorkSheet("Sheet1"); // Access by sheet name
```

or

```cs
// Access Sheet by Index
WorkSheet ws = wb.WorkSheets[0]; // Access by sheet index
```

or

```cs
WorkSheet ws = wb.DefaultWorkSheet; // Access the default sheet
```

or

```cs
WorkSheet ws = wb.WorkSheets.First(); // Access the first sheet
```

or

```cs
WorkSheet ws = wb.WorkSheets.FirstOrDefault(); // Access the first or default sheet
```

Once you have selected the `ws` Worksheet, you can retrieve data from it to perform all Excel-related functions.

<hr class="separator">

## 4. Extract Data from a Worksheet

Data can be retrieved from the `ws` Worksheet as follows:

```cs
string content = ws["cell address"].ToString(); // Extract as a string
int value = ws["cell address"].Int32Value; // Extract as an integer
```

This allows retrieval of data from multiple cells in a specified column:

```cs
foreach (var cell in ws["A2:A10"])
{
    Console.WriteLine("Value is: {0}", cell.Text);
}
```

This code displays the values from cells `A2` to `A10`.

Below is a comprehensive example demonstrating the above concepts:

```cs
// Access Worksheet Data
using IronXL;

static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    foreach (var cell in ws["A2:A10"])
    {
        Console.WriteLine("Value is: {0}", cell.Text);
    }
    Console.ReadKey();
}
```

The result will look like this:

<center>
    <div class="center-image-wrapper">
        <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-input1.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-input1.png" alt="" class="img-responsive add-shadow"></a>
    </div>
</center>

With the `Sample.xlsx` Excel file:

<center>
    <div class="center-image-wrapper">
        <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-1.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-1.png" alt="" class="img-responsive add-shadow"></a>
    </div>
</center>

You'll see how simple it is to manipulate Excel file data in your C# projects using these methods.

<hr class="separator">

## 5. Apply Functions to Data

It's straightforward to manipulate data in an Excel Worksheet by applying functions like Sum, Min, and Max:

```cs
decimal sum = ws["From:To"].Sum();
decimal min = ws["From:To"].Min();
decimal max = ws["From:To"].Max();
```

For more comprehensive information, view our detailed guide on [Writing C# Excel Files](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/#advanced-operations-sum-avg-count-etc) covering aggregate functions.

```cs
// Sum Min Max Functions
using IronXL;

static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");

    decimal sum = ws["G2:G10"].Sum();
    decimal min = ws["G2:G10"].Min();
    decimal max = ws["G2:G10"].Max();

    Console.WriteLine("Sum is: {0}", sum);
    Console.WriteLine("Min is: {0}", min);
    Console.WriteLine("Max is: {0}", max);
    Console.ReadKey();
}
```

This code will produce the following output:

<center>
    <div class="center-image-wrapper">
        <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-output2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc3-output2.png" alt="" class="img-responsive add-shadow"></a>
    </div>
</center>

And this is how the `Sample.xlsx` file appears:

<center>
    <div class="center-image-wrapper">
        <a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc10-2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-read-xlsx-file/doc10-2.png" alt="" class="img-responsive add-shadow"></a>
    </div>
</center>

<hr class="separator">

## 6. Convert a Worksheet to a DataTable

Using IronXL allows straightforward manipulation of an Excel Worksheet as a DataTable.

```cs
DataTable dt = ws.ToDataTable();
```

If you want the first row of the Worksheet to serve as DataTable column names:

```cs
DataTable dt = ws.ToDataTable(true);
```

The Boolean parameter of `ToDataTable()` determines whether the first row should be used as column names in your DataTable. By default, its value is `False`.

```cs
// Convert WorkSheet to DataTable
using IronXL;
using System.Data;

static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    DataTable dt = ws.ToDataTable(true); // Convert sheet1 of sample.xlsx to a datatable
    foreach (DataRow row in dt.Rows) // Access rows
    {