# Saving and Exporting Excel Spreadsheets

The `DataSet` class, part of Microsoft's .NET framework, is a core element of ADO.NET technology, predominantly used in database-related applications. It facilitates interaction with data stemming from various sources, including databases, XML, and more.

IronXL empowers users to transform Excel documents into several formats and inline code objects. The supported file formats include XLS, XLSX, XLSM, CSV, TSV, JSON, XML, and HTML, while inline code objects allow exports as HTML strings, binaries, byte arrays, datasets, and streams.

## Export Spreadsheet Example

Once modifications to an Excel workbook are completed, utilize the `SaveAs` method to convert the workbook to the required file format. This method supports a range of formats like XLS, XLSX, XLSM, CSV, TSV, JSON, XML, and HTML.

Remember to specify the file extension during the import/export process. By default, Excel files are stored in the 'bin > Debug > net6.0' directory within your project.

```cs
using IronXL;

// Instantiate a new Excel workbook
WorkBook workBook = WorkBook.Create();

// Add a new worksheet
WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

// Export the Excel workbook to a specified format
workBook.SaveAs("sample.xls");
```

<hr>

## Exporting CSV, JSON, XML, and HTML Files

While the `SaveAs` method is capable of exporting to formats like CSV, JSON, XML, and HTML, it is preferable to use dedicated methods for these tasks like `SaveAsCsv`, `SaveAsJson`, `SaveAsXml`, and `ExportToHtml`.

```cs
using IronXL;

// Generate a new Excel workbook
WorkBook workBook = WorkBook.Create();

// Add multiple worksheets
WorkSheet workSheet1 = workBook.CreateWorkSheet("sheet1");
WorkSheet workSheet2 = workBook.CreateWorkSheet("sheet2");

// Populate cells
workSheet1["A1"].StringValue = "A1";
workSheet2["A1"].StringValue = "A1";

// Perform exports to various formats
workBook.SaveAsCsv("sample.csv");
workBook.SaveAsJson("sample.json");
workBook.SaveAsXml("sample.xml");
workBook.ExportToHtml("sample.html");
```

It is important to note that for formats like CSV, TSV, JSON, and XML, a unique file is generated for each sheet following the naming pattern **fileName.sheetName.format**. Here is an example of the naming convention for different file formats:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/export-spreadsheet/naming-convention.webp" alt="Naming format" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Exporting to Inline Code Objects

Convert your Excel workbooks to various inline code objects like HTML strings, binary data, byte arrays, streams, and data sets. These exported objects are immediately usable for further application processing.

```cs
using IronXL;
using System.IO;

// Instantiate a new Excel workbook
WorkBook workBook = WorkBook.Create();

// Add a new worksheet
WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

// Convert to HTML string
string htmlString = workBook.ExportToHtmlString();

// Convert to binary and byte arrays
byte[] binary = workBook.ToBinary();
byte[] byteArray = workBook.ToByteArray();

// Convert to stream
Stream stream = workBook.ToStream();

// Convert to DataSet for seamless integration with DataGrids, and SQL
System.Data.DataSet dataSet = workBook.ToDataSet();
```