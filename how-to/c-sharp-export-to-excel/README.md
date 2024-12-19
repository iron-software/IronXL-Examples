# C# Export to Excel: Step by Step Guide

***Based on <https://ironsoftware.com/how-to/c-sharp-export-to-excel/>***


Working with Excel spreadsheets in various formats is a common requirement in many projects, and the ability to export data from these spreadsheets using C# is a critical skill. Whether you're dealing with formats like `.xml`, `.csv`, `.xls`, `.xlsx`, or `.json`, this guide will show you how to proficiently export Excel spreadsheet data using C#. This tutorial will demonstrate a straightforward method that does not rely on the older Microsoft.Office.Interop.Excel library.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>C# Export to Excel</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-get-the-ironxl-library">Acquire the IronXL Library</a></li>
        <li><a href="#anchor-3-c-num-export-to-xlsx-file">Exporting to XLSX using C#</a></li>
        <li><a href="#anchor-5-c-num-export-to-csv-file">Export to CSV in C#</a></li>
        <li><a href="#anchor-6-c-num-export-to-xml-file">Export to XML from Excel in C#</a></li>
        <li><a href="#anchor-7-c-num-export-to-json-file">Handling JSON and Other Formats</a></li>
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

## 1. Acquire the IronXL Library 

IronXL offers a streamlined solution for managing Excel files in .NET Core. You can [download the IronXL DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.export.excel.zip) or [install it via NuGet](https://www.nuget.org/packages/IronXL.Excel) to start using it in your development projects at no cost.

```shell
Install-Package IronXL.Excel
```

After downloading, ensure you reference it in your project, making the `IronXL` classes available through the `IronXL` namespace.

<hr class="separator">
<h4 class="tutorial-segment-title">Tutorial Segment</h4>

## 2. Exporting Data to Excel in C#

IronXL is the simplest tool for exporting data to `.xls`, `.xlsx`, and `.csv` file formats in .NET applications. Additionally, it supports exporting into `.json` and `.xml` formats. Explore how straightforward it is to convert Excel file data into these formats.

<hr class="separator">

## 3. Export to .XLSX File

Exporting to an `.xlsx` file format is particularly straightforward. Here’s an example where we have an existing `XlsFile.xls` in our project's `bin>Debug` directory.

Always remember to include the file extension when importing or exporting.

Normally, new Excel files are saved in the project's `bin>Debug` directory. To save to a custom path, use `wb.SaveAs(@"E:\IronXL\NewXlsxFile.xlsx");`. Learn more on [exporting Excel files in .NET](https://ironsoftware.com/csharp/excel/#convert-excel-spreadsheet).

```cs
/**
Export to XLSX
anchor-c-export-to-xlsx-file
**/
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("XlsFile.xls");//Import .xls, .csv, or .tsv file
    wb.SaveAs("NewXlsxFile.xlsx");//Export as .xlsx file
}
```
<hr class="separator">

## 4. Export to .XLS File

Similarly, exporting to an `.xls` format is also feasible using IronXL. The following example demonstrates this process.

```cs
/**
Export to XLS
anchor-c-export-to-xls-file
**/
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("XlsxFile.xlsx");//Import .xlsx, .csv or .tsv file
    wb.SaveAs("NewXlsFile.xls");//Export as .xls file
}
```
<hr class="separator">

## 5. Export to .CSV File

Here's how you can convert your `.xlsx` or `.xls` files into `.csv` format using IronXL. The example highlights an approach for exporting Excel data to a CSV file.

```cs
/**
Export to CSV
anchor-c-export-to-csv-file
**/
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");  //Import .xlsx or .xls file          
    wb.SaveAsCsv("NewCsvFile.csv"); //Export as .csv file
}
```