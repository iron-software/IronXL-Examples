# C# Export to Excel: Comprehensive Guide

Handling various Excel spreadsheet formats, such as `.xml`, `.csv`, `.xls`, `.xlsx`, and `.json`, is a frequent necessity in programming projects. This guide will help you explore the methods to export data from Excel spreadsheets into these different formats using C#. This process is straightforward and does not require the older Microsoft.Office.Interop.Excel library.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>C# Export to Excel</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-get-the-ironxl-library">Acquire the IronXL Library</a></li>
        <li><a href="#anchor-3-c-num-export-to-xlsx-file">C# Export to XLSX</a></li>
        <li><a href="#anchor-5-c-num-export-to-csv-file">Export to CSV in C#</a></li>
        <li><a href="#anchor-6-c-num-export-to-xml-file">C# Export to XML</a></li>
        <li><a href="#anchor-7-c-num-export-to-json-file">Export JSON Formats and Beyond</a></li>
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

To efficiently manage Excel files in .NET Core, consider using IronXL. [Download IronXL DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.export.excel.zip) or [install via NuGet](https://www.nuget.org/packages/IronXL.Excel). These options are available at no cost for development use.

```shell
Install-Package IronXL.Excel
```

After downloading, include its reference within your project, making the classes under the `IronXL` namespace accessible.

<hr class="separator">
<h4 class="tutorial-segment-title">How to Tutorial</h4>

## 2. Exporting Excel in C#

IronXL offers the simplest approach for exporting data to Excel files such as `.xls`, `.xlsx`, and `.csv` within .NET applications and extends support for `.json` and `.xml` formats. Let's go through the methods one by one.

<hr class="separator">

## 3. C# Export to .XLSX File

Exporting to an `.xlsx` extension is straightforward. Below is an example where our `XlsFile.xls` resides inside the `bin>Debug` folder of the project.

Do note to always include the file extension when dealing with imports and exports.

New Excel files default to the `bin>Debug` folder, but you can specify custom paths like `wb.SaveAs(@"E:\\IronXL\\NewXlsxFile.xlsx");`. Further details can be found [here on exporting Excel files in .NET](https://ironsoftware.com/csharp/excel/#convert-excel-spreadsheet).

```cs
// Export to XLSX
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("XlsFile.xls"); // Import .xls, .csv, or .tsv file
    wb.SaveAs("NewXlsxFile.xlsx"); // Export as .xlsx file
}
```
<hr class="separator">

## 4. C# Export to .XLS File

Similarly, exporting to a .xls extension is also achievable via IronXL. See the following example:

```cs
// Export to XLS
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("XlsxFile.xlsx"); // Import .xlsx, .csv, or .tsv file
    wb.SaveAs("NewXlsFile.xls"); // Export as .xls file
}
```
<hr class="separator">

## 5. C# Export to .CSV File

Exporting an `.xlsx` or `.xls` file to `.csv` is seamless with IronXL. Below is an example showcasing the process:

```cs
// Export to CSV
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx"); // Import .xlsx or .xls file          
    wb.SaveAsCsv("NewCsvFile.csv"); // Export as .csv file
}
```
The provided code will generate three CSV files:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-export-to-excel/doc2-2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-export-to-excel/doc2-2.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

Explaining the occurrence of three `.csv` files: `sample.xlsx` contained three Worksheets, each of which was exported to a separate `.csv`.

The worksheets in `sample.xlsx` are visualized here:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-export-to-excel/doc2-1.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-export-to-excel/doc2-1.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

If there's only one worksheet, a single `.csv` file will be output.

<hr class="separator">

## 6. C# Export to .XML File

Exporting Excel data to an `.XML` format is facilitated through the following code snippet, which again generates multiple files based on worksheet count:

```cs
// Export to XML
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx"); // Import workbook
    wb.SaveAs("NewXmlFile.xml"); // Export as .xml file
}
```
<hr class="separator">

## 7. C# Export to .JSON File

IronXL simplifies the process of exporting Excel data into JSON format, as shown below. This also results in multiple JSON files if there are multiple worksheets:

```cs
// Export to JSON
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx"); // Import workbook
    wb.SaveAsJson("NewjsonFile.json"); // Export as JSON file
}
```
<hr class="separator">

<h4 class="tutorial-segment-title">Tutorial Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100%; height: 140px;" src="https://ironsoftware.com/img/svgs/documentation.svg" class="img-responsive add-shadow" alt="">
      </div>
    </div>
    <div class="col-sm-8">
      <h3>API Reference</h3>
      <p>Explore the IronXL Documentation which includes detailed information on namespaces, feature sets, methods, classes, and enums.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Read API Reference<i class="fa fa-chevron-right"></i></a>
    </div>
  </div>
</div>