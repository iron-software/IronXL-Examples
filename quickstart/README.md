# Excel Spreadsheet Manipulation in C# & VB.NET with IronXL

IronXL, a product from Iron Software, streamlines the process of dealing with Excel (XLS, XLSX, and CSV) files in C# and other .NET languages without needing Excel installed on the server, nor does it require Interop. It offers an API that is more straightforward and faster than **Microsoft.Office.Interop.Excel.**

IronXL is compatible with:

* .NET Framework 4.6.2 or higher for Windows and Azure
* .NET Core 2.0 or higher, supporting Windows, Linux, MacOS, and Azure
* .NET 5, .NET 6, .NET 7, .NET 8, as well as Mono, Mobile, and Xamarin environments

## Installation of IronXL

Install IronXL by using our NuGet package or by [downloading the DLL directly](https://ironsoftware.com/csharp/excel/packages/IronXL.zip). You will find IronXL's classes under the `IronXL` namespace.

For easy installation, use the Visual Studio NuGet Package Manager:
The package is named **IronXL.Excel**.

```shell
Install-Package IronXL.Excel
```

[NuGet package for IronXL.Excel](https://www.nuget.org/packages/ironxl.excel/) can be accessed directly.

## How to Read from an Excel Document

Here's how you can read Excel sheets using IronXL with a few lines of code:

```cs
using IronXL;

// IronXL supports spreadsheet formats such as: XLSX, XLS, CSV, TSV for reading
WorkBook workBook = WorkBook.Load("data.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Easy selection of cells using Excel notation and retrieving various types of data
int cellValue = workSheet["A2"].IntValue;

// Efficiently read ranges of cells
foreach (var cell in workSheet["A2:B10"])
{
    Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
}
```

## How to Create Excel Documents

IronXL simplifies the creation of new Excel documents in C# or VB.NET:

```cs
using IronXL;

// Instantiate a new Excel WorkBook.
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Author = "IronXL Developer";

// Add a new WorkSheet
WorkSheet workSheet = workBook.CreateWorkSheet("main_sheet");

// Add data and styles to your worksheet
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;

// Save the document
workBook.SaveAs("CreatedExcelFile.xlsx");
```

## Export Options

You can export your document to several formats including CSV, XLS, XLSX, JSON, and XML.

```cs
// Fluent saving options in multiple formats
workSheet.SaveAs("CreatedExcelFile.xls");
workSheet.SaveAs("CreatedExcelFile.xlsx");
workSheet.SaveAsCsv("CreatedExcelFile.csv");
workSheet.SaveAsJson("CreatedExcelFile.json");
workSheet.SaveAsXml("CreatedExcelFile.xml");
```

## Cell and Range Styling

Cell and range styling is possible using IronXL's `Range.Style` object:

```cs
// Apply values and styles to cells
workSheet["A1"].Value = "Hello World";
// Multiple styling attributes can be chained
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600").Type = IronXL.Styles.BorderType.Double;
```

## Sorting Cell Ranges

Sorting functionalities within IronXL allow sorting of Excel cells:

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("data.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Obtain a range from the worksheet
Range range = workSheet["A2:A8"];

// Perform ascending sort on the range
range.SortAscending();
// Save changes
workBook.Save();
```

## Modifying Formulas

It’s just as straightforward to edit formulas as setting a value:

```cs
// Assigning a formula to a cell
workSheet["A1"].Value = "=SUM(A2:A10)";

// Access resulting calculation immediately
decimal sum = workSheet["A1"].DecimalValue;
```

## Why IronXL Stands Out?

IronXL offers an accessible API for .NET developers to handle Excel files. It eliminates the need for installing Excel or using Excel Interop, which simplifies managing spreadsheet files.

## Next Steps

For comprehensive understanding and advanced features, be sure to explore the documentation available in our [.NET API Reference](https://ironsoftware.com/csharp/excel/object-reference/) formatted like MSDN.