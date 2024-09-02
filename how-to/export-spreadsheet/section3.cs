using IronXL;
using System.IO;

// Create new Excel WorkBook document
WorkBook workBook = WorkBook.Create();

// Create a blank WorkSheet
WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

// Export to HTML string
string htmlString = workBook.ExportToHtmlString();

// Export as Byte array
byte[] binary = workBook.ToBinary();
byte[] byteArray = workBook.ToByteArray();

// Export as Stream
Stream stream = workBook.ToStream();

// Export as DataSet
System.Data.DataSet dataSet = workBook.ToDataSet(); // Allow easy integration with DataGrids, SQL and EF