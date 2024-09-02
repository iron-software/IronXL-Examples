# Spreadsheet File Conversion Guide

## Introduction
IronXL provides functionalities to switch between various spreadsheet file formats such as XLS, XLSX, XLSM, XLTX, CSV, TSV, JSON, XML, and HTML. Additionally, it supports various in-line code data formats including HTML strings, Binary data, Byte arrays, Data sets, and Memory streams. You can load a spreadsheet using the `Load` method and convert it to the preferred format using the `SaveAs` method. Learn more about exporting your spreadsheets [here](https://ironsoftware.com/csharp/excel/how-to/c-sharp-export-to-excel/).

## Conversion Example

To transform a spreadsheet file type, initiate by loading a compatible file format and then utilize the `SaveAs` method or other specialized methods from IronXL to achieve accurate data remodeling.

For optimal conversion to formats like CSV, JSON, XML, and HTML, it is suggested to use their respective dedicated methods:

- `SaveAsCsv`
- `SaveAsJson`
- `SaveAsXml`
- `ExportToHtml`

Conversions will create a new file for each worksheet named using the pattern **fileName.sheetName.format**. For instance, the CSV output will appear as **sample.new_sheet.csv**.

```cs
using IronXL;

// Load the workbook from a file
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Convert and save the workbook in various formats
workBook.SaveAs("sample.xls");
workBook.SaveAs("sample.tsv");
workBook.SaveAsCsv("sample.csv");
workBook.SaveAsJson("sample.json");
workBook.SaveAsXml("sample.xml");

// Additional export to HTML
workBook.ExportToHtml("sample.html");
```

## Advanced Usage

Beyond basic file types, IronXL captures a myriad of formats for both ingestion and output.

**To load:**
- XLS, XLSX, XLSM, XLTX
- CSV
- TSV

**To export:**
- XLS, XLSX, XLSM
- CSV, TSV
- JSON
- XML
- HTML
- Inline code data types:
  - HTML string
  - Binary and Byte array
  - Data sets: Seamlessly integrate spreadsheets with `System.Data.DataSet` and `System.Data.DataTable` to link with DataGrids, SQL, and EF databases.
  - Memory stream

These formats make it easy to use the data in a web service API or even convert them directly to PDF documents using IronPDF.

```cs
using IronXL;
using System.IO;

// Load workbook from the file
WorkBook workBook = WorkBook.Load("sample.xlsx");

// Perform conversion and save in multiple formats
workBook.SaveAs("sample.xls");
workBook.SaveAs("sample.xlsx");
workBook.SaveAs("sample.tsv");
workBook.SaveAsCsv("sample.csv");
workBook.SaveAsJson("sample.json");
workBook.SaveAsXml("sample.xml");

// Export to HTML and extract the HTML string
workBook.ExportToHtml("sample.html");
string htmlString = workBook.ExportToHtmlString();

// Convert to Byte array, Binary data, and Data set, also create a Memory Stream
byte[] binaryContent = workBook.ToBinary();
byte[] byteArrayContent = workBook.ToByteArray();
System.Data.DataSet dataSet = workBook.ToDataSet(); // Integration with DataGrids, SQL, and EF
Stream memStream = workBook.ToStream();
```

### Converting a Specific Spreadsheet

#### Visual Representation of the Original Spreadsheet:
<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-xlsx.png" alt="XLSX file" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        XLSX file illustration
    </span>
</div>

#### Outputs from Conversion:

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-tsv.png" alt="sample.Data.tsv" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        TSV File Export
    </span>
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-csv.png" alt="sample.Data.csv" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        CSV File Export
    </span>
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-json.png" alt="sample.Data.json" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        JSON File Export
    </span>
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-xml.png" alt="sample.Data.xml" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        XML File Export
    </span>
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-html.png" alt="sample.html" class="img-responsive add-shadow">
    </div>
	<span class="image-description-text_italic">
        HTML File Export
    </span>
</div>

This comprehensive guide details the process of converting different spreadsheet formats using IronXL, ensuring you harness the full potential of spreadsheet management and transformation.