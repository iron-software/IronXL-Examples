using IronXL;
using System.Data;

// Create dataset
DataSet dataSet = new DataSet();

// Create workbook
WorkBook workBook = WorkBook.Create();

// Load DataSet
WorkBook.LoadWorkSheetsFromDataSet(dataSet, workBook);