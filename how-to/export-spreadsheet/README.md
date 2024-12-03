# How to Save or Export Spreadsheets

***Based on <https://ironsoftware.com/how-to/export-spreadsheet/>***


The `DataSet` class is a pivotal aspect of the ADO.NET architecture within Microsoft’s .NET framework. It facilitates manipulation and access to data originating from various sources such as databases, XML files, and other data types, making it essential for database-related applications.

With IronXL, users can transform Excel workbooks into numerous file formats or into inline code objects for increased flexibility. Supported file formats include `XLS`, `XLSX`, `XLSM`, `CSV`, `TSV`, `JSON`, `XML`, and `HTML`. The inline code objects enable the exporting of an Excel file as a `HTML` string, binary data, byte array, dataset, or stream.

## Export Spreadsheet Tutorial

Once you have completed modifying or reviewing your workbook, you can employ the `SaveAs` method to save your Excel spreadsheet into one of the supported file formats, including `XLS`, `XLSX`, `XLSM`, `CSV`, `TSV`, `JSON`, `XML`, and `HTML`.

It is important to specify the file extension during export. Newly created Excel files will generally be saved in the 'bin > Debug > net6.0' directory within your project.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.ExportSpreadsheet
{
    public class ExportExample
    {
        public void Execute()
        {
            // Instantiate a new Excel workbook
            WorkBook workbook = WorkBook.Create();

            // Generate a new WorkSheet
            WorkSheet sheet = workbook.CreateWorkSheet("example_sheet");

            // Save the document in different formats like XLS, XLSX, XLSM, CSV, TSV, JSON, XML, HTML
            workbook.SaveAs("example.xls");
        }
    }
}
```

<hr>

## Exporting CSV, JSON, XML, and HTML Files

While the `SaveAs` method allows for exporting to `CSV`, `JSON`, `XML`, and `HTML` formats, it is advised to use specific methods tailored for each format. Use `SaveAsCsv`, `SaveAsJson`, `SaveAsXml`, and `ExportToHtml` for best results.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.ExportSpreadsheet
{
    public class ExportFormats
    {
        public void Execute()
        {
            // Instantiate a new Excel WorkBook
            WorkBook workbook = WorkBook.Create();

            // Initialize multiple WorkSheets
            WorkSheet sheet1 = workbook.CreateWorkSheet("worksheetA");
            WorkSheet sheet2 = workbook.CreateWorkSheet("worksheetB");

            // Populate cells with data
            sheet1["A1"].StringValue = "Data1";
            sheet2["A1"].StringValue = "Data2";

            // Export as different file formats
            workbook.SaveAsCsv("data.csv");
            workbook.SaveAsJson("data.json");
            workbook.SaveAsXml("data.xml");
            workbook.ExportToHtml("data.html");
        }
    }
}
```

Please note that for `CSV`, `TSV`, `JSON`, and `XML` file formats, distinct files will be generated for each sheet following the format **fileName.sheetName.format**. The illustration below shows files created for each format.

<div class="content-img-align-center">
    <div class="center-image-wrapper">
         <img src="https://ironsoftware.com/static-assets/excel/how-to/export-spreadsheet/naming-convention.webp" alt="Naming format" class="img-responsive add-shadow">
    </div>
</div>

<hr>

## Export to Inline Code Object

Export the Excel workbook to a variety of inline code objects like HTML strings, binary data, byte arrays, streams, and a .NET `DataSet`. These objects can be directly utilized for further applications.

```cs
using System.IO;
using IronXL.Excel;
namespace ironxl.ExportSpreadsheet
{
    public class InlineExport
    {
        public void Execute()
        {
            // Create a new Excel WorkBook
            WorkBook workbook = WorkBook.Create();

            // Add a fresh WorkSheet
            WorkSheet sheet = workbook.CreateWorkSheet("new_sheet");

            // Export to HTML string
            string htmlContent = workbook.ExportToHtmlString();

            // Export as both binary and byte array
            byte[] binaryData = workbook.ToBinary();
            byte[] byteArrayData = workbook.ToByteArray();

            // Export as Stream
            Stream fileStream = workbook.ToStream();

            // Convert to DataSet for easy data manipulation
            System.Data.DataSet dataSet = workbook.ToDataSet();
        }
    }
}
```
This rephrased content keeps professional and conversational tones, enhancing clarity and comprehension.