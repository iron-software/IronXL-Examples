# Importing and Exporting Data as DataSet

***Based on <https://ironsoftware.com/how-to/export-dataset-datatable/>***


A DataSet in .NET provides an in-memory data store that can hold multiple related tables, along with relationships and constraints—making it versatile for handling data from various sources such as databases, XML files, and more.

Each DataSet comprises one or more DataTables, which are essential components serving as individual tables with rows and columns, similar to database tables. DataTables are used for organizing and handling data in a structured format.

IronXL provides functionalities to import data from a DataSet into a spreadsheet and vice versa, streamlining the transition between in-memory data and spreadsheet formats.

## Loading Data into a Spreadsheet

Begin by using the `WorkBook.LoadWorkSheetsFromDataSet` static method to load a DataSet into a workbook. Ensure that a `WorkBook` instance is already created using the `WorkBook.Create` method. The following code snippet illustrates this process, where a new DataSet and WorkBook are initiated and the DataSet is then loaded into the workbook.

```cs
using System.Data;
using IronXL.Excel;

namespace ExportDatasetDatatable
{
    public class ImportSection
    {
        public void Execute()
        {
            // Initialize a new DataSet
            DataSet dataSet = new DataSet();
            
            // Create a new workbook
            WorkBook workBook = WorkBook.Create();
            
            // Import DataSet into the workbook
            WorkBook.LoadWorkSheetsFromDataSet(dataSet, workBook);
        }
    }
}
```

To deepen your understanding of loading spreadsheets from different sources, explore the [How to Load Existing Spreadsheets](https://ironsoftware.com/csharp/excel/how-to/load-spreadsheet/) article.

<hr>

## Converting Workbook to DataSet

The `ToDataSet` method from IronXL permits the conversion of a workbook back to a DataSet. During this conversion, each worksheet in the workbook is transformed into a DataTable in the resulting DataSet. You can decide whether the first row should serve as the column names using the `useFirstRowAsColumnNames` parameter. Here’s a code example:

```cs
using System.Data;
using IronXL.Excel;

namespace ExportDatasetDatatable
{
    public class ExportSection
    {
        public void Execute()
        {
            // Instantiate a new Excel workbook
            WorkBook workBook = WorkBook.Create();
            
            // Add a new worksheet
            WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");
            
            // Convert the workbook to a DataSet
            DataSet dataSet = workBook.ToDataSet();
        }
    }
}
```

For further insight into exporting data to various file formats, refer to the [How to Save or Export Spreadsheets](https://ironsoftware.com/csharp/excel/how-to/export-spreadsheet/) article.