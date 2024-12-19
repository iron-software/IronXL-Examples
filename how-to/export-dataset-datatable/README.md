# How to Import and Export as DataSet

***Based on <https://ironsoftware.com/how-to/export-dataset-datatable/>***


A DataSet serves as a powerful in-memory data structure that can accommodate numerous related tables, along with their relationships and constraints. It's particularly useful for managing data from diverse sources like databases, XML files, and more.

Central to any DataSet is the DataTable, which functions as a single table composed of rows and columns. It mirrors a typical database table and is essential for organizing and processing data in a structured format.

Using IronXL, you can seamlessly convert a DataSet into a spreadsheet object and later revert it back to a DataSet form.

### Begin with IronXL

## Importing a DataSet

To import a DataSet into a spreadsheet object, leverage the static `LoadWorkSheetsFromDataSet` method from the `WorkBook` class. This method requires an initialized `DataSet` and `WorkBook` instance. Initialize your workbook or spreadsheet first by invoking the `Create` method. Below is an example illustrating how to import a DataSet into a workbook, employing both the DataSet and workbook instances.

```cs
using IronXL;
using System.Data;

// Initialize a new DataSet
DataSet dataSet = new DataSet();

// Initialize a new WorkBook
WorkBook workBook = WorkBook.Create();

// Import DataSet into the WorkBook
WorkBook.LoadWorkSheetsFromDataSet(dataSet, workBook);
```

Learn more about loading spreadsheets from different formats in this [How to Load Existing Spreadsheets](https://ironsoftware.com/csharp/excel/how-to/load-spreadsheet/) guide.

## Exporting a DataSet

Convert a workbook into a `System.Data.DataSet` through the `ToDataSet` method, where each worksheet is represented as a `System.Data.DataTable`. This method is used on your working Excel workbook to transform it into a DataSet object. The `useFirstRowAsColumnNames` parameter helps determine if the first row of the worksheet should be treated as column headings.

```cs
using IronXL;
using System.Data;

// Initialize a new Excel Workbook
WorkBook workBook = WorkBook.Create();

// Add a new Worksheet
WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

// Convert Workbook to DataSet
DataSet dataSet = workBook.ToDataSet();
```

For more insights into exporting spreadsheets to various file formats, take a look at this [How to Save or Export Spreadsheets](https://ironsoftware.com/csharp/excel/how-to/export-spreadsheet/) article.