# How to Import Spreadsheet Data

Comma-Separated Values (CSV) files are widely used for their simplicity in storing tabular data where each value is separated by a comma. Conversely, Tab-Separated Values (TSV) files use tabs as delimiters, which are especially useful when the data itself contains commas.

The `DataSet` class, part of Microsoft .NET's ADO.NET framework, provides the functionality to manage data from various sources, including databases and XML. This makes it invaluable for applications that interact with different kinds of data storage.

Excel file formats such as XLSX, XLS, XLSM, XLTX, and the previously mentioned CSV and TSV, along with `DataSet` objects, can all be imported into an Excel spreadsheet using IronXL.


## Example: Loading Spreadsheets

You can import an existing Excel workbook using IronXL's `WorkBook.Load` method. This method supports multiple file formats including XLSX, XLS, XLSM, XLTX, CSV, and TSV. If your workbook is encrypted with a password, you can supply this password as an additional parameter. Moreover, for working directly with raw data, IronXL provides `FromByteArray` and `FromStream` methods to handle byte arrays and streams respectively.

```cs
using IronXL;

// Load Excel file formats and CSV, TSV
WorkBook workBook = WorkBook.Load("sample.xlsx");
```

## Working with CSV Files

Although the `Load` method is capable of handling various formats, using the `LoadCSV` method is preferable for loading CSV files due to its specific optimizations for handling comma-separated data.

```cs
using IronXL;

// Directly loading a CSV file
WorkBook workBook = WorkBook.LoadCSV("sample.csv");
```

## Importing DataSets

The `DataSet` class in Microsoft .NET can be used as a robust solution for handling in-memory data storage without continuous database connection. You can also integrate a `DataSet` into an Excel workbook using IronXL's `LoadWorkSheetsFromDataSet` method. Below is an example where we first create an empty `DataSet` and then incorporate it into an Excel workbook.

```cs
using IronXL;
using System.Data;

// Initialize an empty DataSet
DataSet dataSet = new DataSet();

// Create a new workbook
WorkBook workBook = WorkBook.Create();

// Import DataSet into the workbook
WorkBook.LoadWorkSheetsFromDataSet(dataSet, workBook);
```
In this approach, `DataSet` can also be populated with data retrieved from a database query, providing a flexible tool for data manipulation and storage within .NET applications.