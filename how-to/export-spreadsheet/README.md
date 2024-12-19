# How to Save or Export Spreadsheets

***Based on <https://ironsoftware.com/how-to/export-spreadsheet/>***


The `DataSet` class, built into Microsoft's .NET framework, is a crucial part of ADO.NET (ActiveX Data Objects for .NET) technology. It is essential for applications dealing with databases and provides the ability to work with data from various sources including databases, XML, etc.

IronXL allows for the conversion of Excel documents into numerous file formats and inline code objects. Supported file formats include XLS, XLSX, XLSM, CSV, TSV, JSON, XML, and HTML. Inline code objects include exporting to HTML strings, binaries, byte arrays, datasets, and streams.

### Getting Started with IronXL

---

## Exporting Spreadsheets

Once you have finished editing or reviewing your workbook, you can use the `SaveAs` method to export the Excel file to a variety of formats such as XLS, XLSX, XLSM, CSV, TSV, JSON, XML, and HTML. Remember to specify the file extension when importing or exporting files. By default, new Excel files are stored in the 'bin > Debug > net6.0' directory of your project.

```cs
using IronXL;

// Instantiate a new Excel WorkBook
WorkBook workBook = WorkBook.Create();

// Add a new WorkSheet
WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

// Demonstrate saving the file in different formats
workBook.SaveAs("sample.xls");
```

---

## Exporting CSV, JSON, XML, and HTML

While the `SaveAs` method supports these file types, you might want to use specialized methods for clarity and functionality, such as `SaveAsCsv`, `SaveAsJson`, `SaveAsXml`, and `ExportToHtml`.

```cs
using IronXL;

// Create a new Excel WorkBook
WorkBook workBook = WorkBook.Create();

// Add multiple WorkSheets
WorkSheet workSheet1 = workBook.CreateWorkSheet("sheet1");
WorkSheet workSheet2 = workBook.CreateWorkSheet("sheet2");

// Populate cells with information
workSheet1["A1"].StringValue = "A1";
workSheet2["A1"].StringValue = "A1";

// Export to different formats
workBook.SaveAsCsv("sample.csv");
workBook.SaveAsJson("sample.json");
workBook.SaveAsXml("sample.xml");
workBook.ExportToHtml("sample.html");
```

Note that for file formats like CSV, TSV, JSON, and XML, individual files are prepared for each sheet following the naming scheme **fileName.sheetName.format**. Below is an illustration depicting this naming schema:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/export-spreadsheet/naming-convention.webp" alt="File Naming Convention" class="img-responsive add-shadow">
    </div>
</div>

---

## Exportation to Inline Code Objects

You can also export Excel workbooks to various inline code objects. These include HTML strings, binary data, byte arrays, streams, and .NET DataSets, allowing for easy integration with such elements as DataGrids, SQL databases, and Entity Framework.

```cs
using IronXL;
using System.IO;

// Initialize a new Excel WorkBook
WorkBook workBook = WorkBook.Create();

// Add a new blank WorkSheet
WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

// Demonstration of various export formats
string htmlString = workBook.ExportToHtmlString();
byte[] binary = workBook.ToBinary();
byte[] byteArray = workBook.ToByteArray();
Stream stream = workBook.ToStream();
System.Data.DataSet dataSet = workBook.ToDataSet(); // Enables seamless integration with DataGrids, SQL databases, and Entity Framework
```