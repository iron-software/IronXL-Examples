# How to Convert Spreadsheet File Formats

***Based on <https://ironsoftware.com/how-to/convert-spreadsheet-file-types/>***


## Introduction
IronXL enables the transformation of spreadsheet files across a variety of formats such as XLS, XLSX, XLSM, XLTX, CSV, TSV, JSON, XML, and HTML. It also accommodates inline code data types like HTML string, Binary, Byte array, Data set, and Memory stream. To open a spreadsheet file, use the `Load` method, and to convert it to a different format, employ the `SaveAs` method. Learn more about exporting spreadsheets with IronXL [here](https://ironsoftware.com/csharp/excel/how-to/c-sharp-export-to-excel/).

***

***

<h3>Introduction to IronXL</h3>

------------------------------

## Example: Converting Types of Spreadsheets

Converting between spreadsheet types with IronXL involves loading a spreadsheet in an accepted format and using the `SaveAs` method to restructure and save it in another format.

While `SaveAs` can handle various formats like CSV, JSON, XML, and HTML, it's better to use format-specific methods for optimized results:

- `SaveAsCsv`
- `SaveAsJson`
- `SaveAsXml`
- `ExportToHtml`

For formats such as CSV, TSV, JSON, and XML, a unique file will be created for each worksheet. The files are named following the pattern **fileName.sheetName.format**. For instance, the output file name for the CSV would be **sample.new_sheet.csv**.

```cs
using IronXL;

// Start by importing any supported spreadsheet format like XLSX, XLS, XLSM, XLTX, CSV, and TSV
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Here's how to save the spreadsheet in different file formats
workBook.SaveAs("sample.xls");
workBook.SaveAs("sample.tsv");
workBook.SaveAsCsv("sample.csv");
workBook.SaveAsJson("sample.json");
workBook.SaveAsXml("sample.xml");

// For converting into an HTML file
workBook.ExportToHtml("sample.html");
```
<hr>

## Advanced Conversion Techniques

While the previous section highlighted commonly used file formats, IronXL supports an even broader array of formats for both loading and exporting data.

### Loading Options
- XLS, XLSX, XLSM, XLTX
- CSV
- TSV

### Exporting Options
- XLS, XLSX, XLSM
- CSV, TSV
- JSON
- XML
- HTML
- Inline code data types:
  - HTML string
  - Binary and Byte array
  - Data set: Converts Excel to `System.Data.DataSet` and `System.Data.DataTable` for seamless integration with DataGrids, SQL, and EF.
  - Memory stream

These data types can also be utilized in RESTful API responses or converted to a PDF using IronPDF.

```cs
using IronXL;
using System.IO;

// Load a spreadsheet in any of the supported formats
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Here's how to export the spreadsheet to various formats
workBook.SaveAs("sample.xls");
workBook.SaveAs("sample.xlsx");
workBook.SaveAs("sample.tsv");
workBook.SaveAsCsv("sample.csv");
workBook.SaveAsJson("sample.json");
workBook.SaveAsXml("sample.xml");

// Converting to HTML and HTML string format
workBook.ExportToHtml("sample.html");
string htmlString = workBook.ExportToHtmlString();

// Exporting as Binary, Byte array, Data set, and Stream
byte[] binary = workBook.ToBinary();
byte[] byteArray = workBook.ToByteArray();
System.Data.DataSet dataSet = workBook.ToDataSet(); // Enables easy integration with DataGrids, SQL, and EF
Stream stream = workBook.ToStream();
```

The code demonstrates loading a typical XLSX file and exporting it in several output formats.

### Sample Spreadsheet for Conversion
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-xlsx.png" alt="XLSX file" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        XLSX file image
    </span>
</div>

Displayed below are the various exported spreadsheet files.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-tsv.png" alt="sample.Data.tsv" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        TSV File Export Image
    </span>
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-csv.png" alt="sample.Data.csv" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        CSV File Export Image
    </span>
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-json.png" alt="sample.Data.json" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        JSON File Export Image
    </span>
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-xml.png" alt="sample.Data.xml" class's="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        XML File Export Image
    </span>
	    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-html.png" alt="sample.html" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        HTML File Export Image
    </span>
</div>