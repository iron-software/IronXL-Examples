using IronXL;
using System.Data;

// Create new Excel WorkBook document
WorkBook workBook = WorkBook.Create();

// Create a blank WorkSheet
WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

// Export as DataSet
DataSet dataSet = workBook.ToDataSet();