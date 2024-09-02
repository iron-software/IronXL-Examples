# How to Import Excel Files in C#

Developers frequently need to pull data from Excel files for use in application development and data management. IronXL offers a streamlined approach to precisely import and programmatically handle the data needed within a C# project, without necessitating extensive coding.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Import Excel Data C#</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-import-data-with-the-ironxl-library">Import Data with the IronXL Library</a></li>
        <li><a href="#anchor-3-import-excel-data-in-c-num">Import Excel data in C#</a></li>
        <li><a href="#anchor-4-import-excel-data-of-specific-range">Import data of specific cell range</a></li>
        <li><a href="#anchor-5-import-excel-data-by-aggregate-functions">Import Excel data with aggregate functions SUM, AVG, MIN, MAX, and more</a></li>
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

## 1. Import Data with the IronXL Library

Using IronXL, we can import data efficiently into our C# project. This library is freely available for development.

Install into your C# project either [via direct DLL Download](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.Import.Excel.Csharp.zip) or by using [the NuGet package manager](https://www.nuget.org/packages/IronXL.Excel).

<br>

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">How to Tutorial</h4>

## 2. Access WorkSheet for Project

Begin by leveraging IronXL, installed in the first step, to import Excel data into our C# development.

To load an Excel WorkBook into our C# project, we utilize the `WorkBook.Load()` method of IronXL, specifying the file path as a string parameter:

```cs
// Load Excel file
WorkBook wb = WorkBook.Load("Path");
```

Next, we select a WorkSheet from this WorkBook for data importing, using the `GetWorkSheet()` method by specifying the sheet name:

```cs
// Specify the sheet name of the Excel WorkBook
WorkSheet ws = wb.GetWorkSheet("SheetName");
```

`ws` now represents the selected WorkSheet, and `wb` is the WorkBook loaded earlier. Here are some alternative methods for incorporating an Excel WorkSheet into your project:

```cs
/**
Import WorkSheet 
anchor-access-worksheet-for-project
**/
// By sheet index
WorkBook.WorkSheets[SheetIndex];
// Get default WorkSheet
WorkBook.DefaultWorkSheet;
// Get first WorkSheet
WorkBook.WorkSheets.First();
// For the first or default sheet
WorkBook.WorkSheets.FirstOrDefault();
```

Now, we are ready to import various types of data from the specified Excel files. Let's explore all possible ways to utilize Excel file data in our project.

<hr class="separator">

## 3. Import Excel Data in C&num;

To import specific data from an Excel file, we use a cell addressing system:

```cs
WorkSheet["Cell Address"];
```

Alternatively, we can access data using the row and column indices:

```cs
WorkSheet.Rows[RowIndex].Columns[ColumnIndex]
```

To store imported cell values into variables:

```cs
/**
Import Data by Cell Address
anchor-import-excel-data-in-c-num
**/
// By cell addressing
string val = WorkSheet["Cell Address"].ToString();
// By row and column indexing
string.val = WorkSheet.Rows[RowIndex].Columns[ColumnIndex].Value.ToString();
```

These examples employ zero-based indexing for rows and columns.

<hr class="separator">

## 4. Import Excel Data of Specific Range

To import data from a defined range in an Excel WorkBook, utilize the `range` method specifying the start and end cell addresses. This will retrieve all cell values within the designated range.

```cs
WorkSheet["starting Cell Address: Ending Cell Address"];
```

Learn more about [working with ranges in Excel files](https://ironsoftware.com/csharp/excel/#excel-ranges) through the provided code examples.

```cs
/**
Import Data by Range
anchor-import-excel-data-of-specific-range
**/
using IronXL;
static void Main(string[] args)
{
    // Import Excel WorkBook
    WorkBook wb = WorkBook.Load("sample.xlsx");
    // Specify WorkSheet
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Import data of specific cell
    string val = ws["A4"].Value.ToString();
    Console.WriteLine("Import Value of A4 Cell address: {0}", val);
    Console.WriteLine("import Values in Range From B3 To B9:\n");
    // Import data in specific range
    foreach (var item in ws["B3:B9"])
    {
        Console.WriteLine(item.Value.ToString());
    }

    Console.ReadKey();
}
```

The code above produces the output:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-import-excel/1output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-import-excel/1output.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

With the values of Excel file `sample.xlsx` as:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-import-excel/1excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-import-excel/1excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## 5. Import Excel Data by Aggregate Functions

Aggregate functions can be applied to Excel files to retrieve computed data. Here are some common functions and their usage:

* `Sum()`
 ```cs
 // Find the sum of a specific cell range
WorkSheet["Starting Cell Address: Ending Cell Address"].Sum();
```

* `Average()`
```cs
 // Calculate the average of a specific cell range
WorkSheet["Starting Cell Address: Ending Cell Address"].Avg()
```

* `Min()`
```cs
 // Get the minimum value in a specific cell range
WorkSheet["Starting Cell Address: Ending Cell Address"].Min()
```

* `Max()`
```cs
 // Get the maximum value in a specific[cell range]
WorkSheet["Starting Cell Address: Ending Cell Address"].Max()
```

For further details, explore the guide on [working with aggregate functions in Excel for C#](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/#advanced-operations-sum-avg-count-etc).

Example of applying these functions to import Excel file data:

```cs
/**
Import Data by Aggregate Function
anchor-import-excel-data-by-aggregate-functions
**/
using IronXL;
static void Main(string [] args)
{
    // Import Excel file
    WorkBook wb = WorkBook.Load("sample.xlsx");
    // Specify WorkSheet
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Apply aggregate functions to import data
    decimal Sum = ws["D2:D9"].Sum();
    decimal Avg = ws["D2:D9"].Avg();
    decimal Min = ws["D2:D9"].Min();
    decimal Max = ws["D2:D9"].Max();
    Console.WriteLine("Sum From D2 To D9: {0}", Sum);
    Console.WriteLine("Avg From D2 To D9: {0}", Avg);
    Console.WriteLine("Min From D2 To D9: {0}", Min);
    Console.WriteLine("Max From D2 To D9: {0}", Max);
    Console.ReadKey();
}
```

This code results in the following output:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-import-excel/2output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-import-excel/2output.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

And the values of `sample.xlsx` in this instance are:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-import-excel/2excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-import-excel/2excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## 6. Import Complete Excel File Data

To import the entirety of an Excel file into a C# project, begin by converting the loaded WorkBook into a DataSet, converting its WorkSheets into DataTables therein:

```cs
// Convert WorkBook into DataSet
DataSet ds = WorkBook.ToDataSet();
```

When importing, if the first column in the Excel file serves as the header:

```cs
// Set the first column as DataTable ColumnName
ToDataSet(true);
```

This code will configure the initial column in the Excel file as the DataTable ColumnNames.

An example on how to import an Excel file into a DataSet and utilize the first Excel sheet column as DataTable ColumnName:

```cs
/**
Import to Dataset
anchor-import-complete-excel-file-data
**/
using IronXL;
static