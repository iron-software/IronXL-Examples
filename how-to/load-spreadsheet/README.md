# Guide to Importing Spreadsheets into IronXL

***Based on <https://ironsoftware.com/how-to/load-spreadsheet/>***


The CSV (Comma-Separated Values) file format is predominantly used for tabular data in which values are segregated by commas. This format is extensively utilized for data interchange. Conversely, the TSV (Tab-Separated Values) format, which employs tabs as delimiters, is preferable when the data includes commas.

Microsoft's .NET incorporates the `DataSet` class as part of its ADO.NET (ActiveX Data Objects for .NET) framework, frequently deployed in database-driven applications. This class facilitates operations with data from various sources, including databases, XML files, and more.

Excel file formats such as XLSX, XLS, XLSM, XLTX, along with CSV and TSV, can be seamlessly integrated into an Excel spreadsheet utilizing IronXL.

## Loading Excel Spreadsheets

To open an existing Excel workbook, leverage the `Load` method provided by IronXL. This method caters to multiple file formats like XLSX, XLS, XLSM, XLTX, CSV, and TSV. Should the workbook be secured with a password, it can be accessed by supplying the password as a secondary argument to the method. Additionally, workbook data can be supplied using a byte array or a stream, facilitated by the `FromByteArray` and `FromStream` methods respectively.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.LoadSpreadsheet
{
    public class Section1
    {
        public void Run()
        {
            // Accepts formats: XLSX, XLS, XLSM, XLTX, CSV, and TSV
            WorkBook workBook = WorkBook.Load("sample.xlsx");
        }
    }
}
```

<hr>

## Importation of CSV Files

While the `Load` method is equipped to handle all supported file formats, it is advisable to employ the `LoadCSV` method specifically for CSV files.

```cs
using IronXL;
using IronXL.Excel;
namespace ironxl.LoadSpreadsheet
{
    public class Section2
    {
        public void Run()
        {
            // Primary method for loading CSV files
            WorkBook workBook = WorkBook.LoadCSV("sample.csv");
        }
    }
}
```

<hr>

## Integrating DataSets

Utilizing the `DataSet` class from Microsoft .NET allows for the management of data in a disconnected, memory-resident layout. These DataSets can also be integrated into an IronXL workbook using the `LoadWorkSheetsFromDataSet` method. Here, an empty DataSet is created for demonstration, although commonly, DataSets are populated from database queries.

```cs
using System.Data;
using IronXL.Excel;
namespace ironxl.LoadSpreadsheet
{
    public class Section3
    {
        public void Run()
        {
            // Initialize an empty DataSet
            DataSet dataSet = new DataSet();
            
            // Generate a new workbook
            WorkBook workBook = WorkBook.Create();
            
            // Embed DataSet into the workbook
            WorkBook.LoadWorkSheetsFromDataSet(dataSet, workBook);
        }
    }
}
```