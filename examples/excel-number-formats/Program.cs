using IronXL;
using System;
using System.Linq;

// Load an existing WorkSheet
WorkBook workBook = WorkBook.Load("sample.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Set data display format to cell
// The cell value will look like 12300%
workSheet["A2"].Value = 123;
workSheet["A2"].FormatString = "0.0%";

// The cell value will look like 123.0000
workSheet["A2"].First().FormatString = "0.0000";

// Set data display format to range
DateTime dateValue = new DateTime(2020, 1, 1, 12, 12, 12);
workSheet["A3"].Value = dateValue;
workSheet["A4"].First().Value = new DateTime(2022, 3, 3, 10, 10, 10);
workSheet["A5"].First().Value = new DateTime(2021, 2, 2, 11, 11, 11);

var range = workSheet["A3:A5"];

// The cell(A3) value will look like 1/1/2020 12:12:12 PM
range.FormatString = "MM/dd/yy h:mm:ss";

workBook.SaveAs("numberFormats.xls");
