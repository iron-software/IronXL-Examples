# How to Load Existing Spreadsheets

***Based on <https://ironsoftware.com/how-to/load-spreadsheet/>***


The CSV (Comma-Separated Values) format is widely utilized for tabular data where each value is separated by a comma, making it ideal for data sharing. In contrast, the TSV (Tab-Separated Values) format employs tabs as separators and is preferred when the data includes commas.

`DataSet`, an integral part of the .NET framework under ADO.NET (ActiveX Data Objects for .NET), is typically utilized in database-related applications. It facilitates interactions with data originating from various sources such as databases and XML files.

These data types, encompassing formats like XLSX, XLS, XLSM, XLTX, CSV, and TSV, can be effectively managed within an Excel spreadsheet through IronXL.

### Starting with IronXL

---

## Spreadsheet Loading Example

To initiate the loading of an existing Excel workbook, employ the `Load` static method. This method is compatible with various formats including XLSX, XLS, XLSM, XLTX, CSV, and TSV. If encountering a protected workbook, a password parameter can be provided. Additionally, the workbook data can be presented via byte arrays or streams, which are handled by `FromByteArray` or `FromStream` methods respectively.

```cs
using IronXL;

// Compatible with XLSX, XLS, XLSM, XLTX, CSV, and TSV formats
WorkBook workBook = WorkBook.Load("sample.xlsx");
```

---

## Loading a CSV File

For specifically handling CSV files, opt for the `LoadCSV` method, despite the general `Load` method's ability to read all supported formats.

```cs
using IronXL;

// Specifically for loading CSV files
WorkBook workBook = WorkBook.LoadCSV("sample.csv");
```

---

## Incorporating a DataSet

In the .NET framework, `DataSet` is employed for the management of data in a disconnected, in-memory format. This data can be loaded into an Excel workbook using IronXL's `LoadWorkSheetsFromDataSet` method. Below is an example where a DataSet is instantiated; it is more typical to fill a DataSet via a database query.

```cs
using IronXL;
using System.Data;

// Instantiate a DataSet
DataSet dataSet = new DataSet();

// Create a new workbook
WorkBook workBook = WorkBook.Create();

// Load the DataSet into the workbook
WorkBook.LoadWorkSheetsFromDataSet(dataSet, workBook);
```