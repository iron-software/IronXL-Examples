using IronXL;
using System.Linq;

WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
WorkSheet workSheet1 = workBook.CreateWorkSheet("Sheet1");
WorkSheet workSheet2 = workBook.CreateWorkSheet("Sheet2");

// Create workbook(global) define name
workSheet1["D5"].SaveAsNamedRange("Iron", true);

// Create worksheet define name
workSheet2["D10"].SaveAsNamedRange("Hello", false);

// --== Within the same worksheet ==--
// Set hyperlink to cell Z20
workSheet1["A1"].Value = "Z20";
workSheet1["A1"].First().Hyperlink = "Z20";

// Set hyperlink to define name "Iron"
workSheet1["A2"].Value = "Iron";
workSheet1["A2"].First().Hyperlink = "Iron";

// --== Across worksheet ==--
// Set hyperlink to cell A1 of Sheet2
workSheet1["A3"].Value = "A1 of Sheet2";
workSheet1["A3"].First().Hyperlink = "Sheet2!A1";

// Set hyperlink to define name "Hello" of Sheet2
workSheet1["A4"].Value = "Define name Hello of Sheet2";
workSheet1["A4"].First().Hyperlink = "Sheet2!Hello";

workBook.SaveAs("setHyperlinkAcrossWorksheet.xlsx");