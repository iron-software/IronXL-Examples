# How to Convert Spreadsheet File Types

***Based on <https://ironsoftware.com/how-to/convert-spreadsheet-file-types/>***


## Introduction
IronXL facilitates the transformation of spreadsheet files across a variety of formats such as XLS, XLSX, XLSM, XLTX, CSV, TSV, JSON, XML, and HTML. It supports different inline code data types including HTML string, Binary, Byte array, Data set, and Memory stream. To open a spreadsheet file, utilize the `Load` method, and to save the spreadsheet in a specific format, employ the `SaveAs` method or follow this [export guide](https://ironsoftware.com/csharp/excel/how-to/c-sharp-export-to-excel/).

***

## Convert Spreadsheet Type Example

The conversion process of spreadsheet files using IronXL involves loading a file in any supported format and saving it in another through the intelligent data restructuring features of IronXL.

Although the `SaveAs` method can be utilized for formats like CSV, JSON, XML, and HTML, it is advisable to use the format-specific methods:

- `SaveAsCsv`
- `SaveAsJson`
- `SaveAsXml`
- `ExportToHtml`

For file formats such as CSV, TSV, JSON, and XML, each worksheet is saved as a separate file, following the pattern **fileName.sheetName.format**. For example, the output for the CSV format would be **sample.new_sheet.csv**.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.ConvertSpreadsheetFileTypes
{
    public class Section1
    {
        public void Run()
        {
            // Load spreadsheets in formats like XLSX, XLS, XLSM, XLTX, CSV, and TSV
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            
            // Save the spreadsheet in formats like XLS, XLSX, XLSM, CSV, TSV, JSON, XML
            workBook.SaveAs("sample.xls");
            workBook.SaveAs("sample.tsv");
            workBook.SaveAsCsv("sample.csv");
            workBook.SaveAsJson("sample.json");
            workBook.SaveAsXml("sample.xml");
            
            // Also save the spreadsheet as an HTML file
            workBook.ExportToHtml("sample.html");
        }
    }
}
```
<hr>

## Advanced

In the previous example, we covered common file formats for conversion. Yet, IronXL's ability to convert spreadsheets encompasses a broader range of formats. Check out all the options available for loading and exporting spreadsheets.

### Load
- XLS, XLSX, XLSM, and XLTX
- CSV
- TSV

### Export
- XLS, XLSX, and XLSM
- CSV and TSV
- JSON
- XML
- HTML
- Inline code data types:
	- HTML string
	- Binary and Byte array
	- Data set: Converts Excel into `System.Data.DataSet` and `System.Data.DataTable` objects facilitating seamless integration with DataGrids, SQL, and EF.
	- Memory stream

These inline code data types can be transmitted as RESTful API responses or utilized with IronPDF to convert them to PDF documents.

```cs
using System.IO;
using IronXL.Excel;
namespace ironxl.ConvertSpreadsheetFileTypes
{
    public class Section2
    {
        public void Run()
        {
            // Load any format like XLSX, XLS, XLSM, XLTX, CSV, and TSV
            WorkBook workBook = WorkBook.Load("sample.xlsx");
            
            // Output the spreadsheet into formats such as XLS, XLSX, XLSM, CSV, TSV, JSON, XML
            workBook.SaveAs("sample.xls");
            workBook.SaveAs("sample.xlsx");
            workBook.SaveAs("sample.tsv");
            workBook.SaveAsCsv("sample.csv");
            workBook.SaveAsJson("sample.json");
            workBook.SaveAsXml("sample.xml");
            
            // Additionally, export the spreadsheet to HTML and HTML string
            workBook.ExportToHtml("sample.html");
            string htmlString = workBook.ExportToHtmlString();
            
            // Transform the spreadsheet to Binary, Byte array, Data set, and Stream formats
            byte[] binary = workBook.ToBinary();
            byte[] byteArray = workBook.ToByteArray();
            System.Data.DataSet dataSet = workBook.ToDataSet();
            Stream stream = workBook.ToStream();
        }
    }
}
```

### The Spreadsheet We Will Convert
![XLSX file](https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-xlsx.png)

The following images illustrate the different files exported:

![TSV File Export](https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-tsv.png)

![CSV File Export](https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-csv.png)

![JSON File Export](https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-json.png)

![XML File Export](https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-xml.png)

![HTML File Export](https://ironsoftware.com/static-assets/excel/how-to/convert-spreadsheet-file-types/convert-spreadsheet-file-types-html.png)