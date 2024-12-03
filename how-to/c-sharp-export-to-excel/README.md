# C# Export to Excel: Comprehensive Guide

***Based on <https://ironsoftware.com/how-to/c-sharp-export-to-excel/>***


Managing various Excel spreadsheet formats and leveraging C# export functionalities is often essential in projects that require manipulating `.xml`, `.csv`, `.xls`, `.xlsx`, and `.json` data. This guide provides a step-by-step process for exporting Excel data into these formats using C#, efficiently and without the need for the outdated Microsoft.Office.Interop.Excel library.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>C# Export to Excel</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-get-the-ironxl-library">Acquire the IronXL Library</a></li>
        <li><a href="#anchor-3-c-num-export-to-xlsx-file">Export to XLSX with C#</a></li>
        <li><a href="#anchor-5-c-num-export-to-csv-file">Export to CSV with C#</a></li>
        <li><a href="#anchor-6-c-num-export-to-xml-file">Export to XML with C#</a></li>
        <li><a href="#anchor-7-c-num-export-to-json-file">Export to JSON and more</a></li>
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

For streamlined handling of Excel files within .NET Core, consider using IronXL. [Download the IronXL DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.export.excel.zip) or [install it via NuGet](https://www.nuget.org/packages/IronXL.Excel) for development purposes at no cost.

```shell
Install-Package IronXL.Excel
```

After downloading, incorporate its reference into your project, and access `IronXL` classes via the `IronXL` namespace.

<hr class="separator">
<h4 class="tutorial-segment-title">Export Guide</h4>

## 2. C# Export Methods

IronXL simplifies data export to Excel and supports exporting to formats like `.xls`, `.xlsx`, and `.csv`. Additionally, it handles data conversion to `.json` and `.xml`. Let's explore how effortlessly Excel file data can be converted to these formats.

<hr class="separator">

## 3. Exporting to .XLSX File

Exporting an Excel file to `.xlsx` format is straightforward. The code below demonstrates this by working with an existing file located in your project's `bin>Debug` folder.

Remember: Always specify the file extension during import and export operations.

For example, to create a new file in a specific directory, use:

```cs
using IronXL;

static void Main(string[] args) {
    WorkBook wb = WorkBook.Load("XlsFile.xls");
    wb.SaveAs(@"E:\IronXL\NewXlsxFile.xlsx");
}
```

Learn more about exporting Excel files in .NET from this [detailed tutorial](https://ironsoftware.com/csharp/excel/#convert-excel-spreadsheet).

<hr class="separator">

## 4. Exporting to .XLS File

Similarly, exporting data to a `.xls` file can be accomplished with IronXL as illustrated below:

```cs
using IronXL;

static void Main(string[] args) {
    WorkBook wb = WorkBook.Load("XlsxFile.xlsx");
    wb.SaveAs("NewXlsFile.xls");
}
```

<hr class="separator">

## 5. Exporting to .CSV File

The following example showcases the conversion of `.xlsx` or `.xls` files to `.csv`. This process is illustrated when the `sample.xlsx` file containing multiple worksheets results in separate `.csv` files for each worksheet.

```cs
using IronXL;

static void Main(string[] args) {
    WorkBook wb = WorkBook.Load("sample.xlsx");
    wb.SaveAsCsv("NewCsvFile.csv");
}
```

The resulting `.csv` files are created for each worksheet within the Excel file. If `sample.xlsx` includes only one worksheet, a single `.csv` file will be generated.

<hr class="separator">

## 6. Exporting to .XML File
 
Exporting Excel data to `.xml` format follows a similar approach, where multiple XML files may result from a single workbook containing multiple sheets.

```cs
using IronXL;

static void Main(string[] args) {
    WorkBook wb = WorkBook.Load("sample.xlsx");
    wb.SaveAsCsv("NewXmlFile.xml");
}
```

<hr class="separator">

## 7. Exporting to .JSON File
 
Finally, converting Excel data to JSON format is made easy with IronXL. Just as with previous formats, this may result in multiple JSON files if the original workbook contains multiple sheets.

```cs
using IronXL;

static void Main(string[] args) {
    WorkBook wb = WorkBook.Load("sample.xlsx");
    wb.SaveAsJson("NewjsonFile.json");
}
```

<hr class="separator">

<h4 class="tutorial-segment-title">Quick Access to Tutorial Resources</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg">
      </div>
    </div>
    <div class="col-sm-8">
      <h3>API Reference</h3>
      <p>Explore comprehensive IronXL Documentation, including all namespaces, methods, properties, classes, and enums.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Read API Reference<i class="fa fa-chevron-right"></i></a>
    </div>
  </div>
</div>