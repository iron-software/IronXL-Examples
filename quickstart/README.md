# Working with Excel Files in C# & VB.NET Applications

***Based on <https://ironsoftware.com/docs/docs/>***


Manipulating Excel files such as XLS, XLSX, and CSV is straightforward in C# and other .NET languages with the help of the IronXL library from Iron Software.

IronXL eliminates the need for having Excel installed on your server or using Interop. It offers an API that is both more efficient and more user-friendly than **Microsoft.Office.Interop.Excel**.

IronXL is compatible with a range of platforms:

* .NET Framework 4.6.2 and newer on Windows and Azure
* .NET Core 2 and newer on Windows, Linux, MacOS, and Azure
* .NET 5, .NET 6, .NET 7, .NET 8, Mono, Mobile, and Xamarin


## Setting Up IronXL

To begin using IronXL, you can either install it via NuGet package or by [downloading the DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip). The `IronXL` namespace encompasses all necessary classes.

The simplest method to integrate IronXL is through the NuGet Package Manager in Visual Studio:
The required package name is **IronXL.Excel**.

```shell
Install-Package IronXL.Excel
```

[NuGet Package Link](https://www.nuget.org/packages/ironxl.excel/)


## How to Read an Excel Document

Fetching data from an Excel file is simple with IronXL and only takes a few lines of code:

```cs
using IronXL;

// Formats supported: XLSX, XLS, CSV, and TSV
WorkBook workBook = WorkBook.Load("data.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Access cells using Excel references to retrieve values, dates, texts, or formulas
int cellValue = workSheet["A2"].IntValue;

// Iterate through a range of cells
foreach (var cell in workSheet["A2:B10"])
{
    Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
}
```

## Creating Excel Documents

IronXL also allows for easy creation of new Excel documents in C# or VB.NET:

```cs
using IronXL;

// Generate a new Excel WorkBook.
WorkBook workBook = WorkBook.Create(ExcelFileFormat.XLSX);
workBook.Metadata.Author = "IronXL";

// Add a new WorkSheet.
WorkSheet workSheet = workBook.CreateWorkSheet("main_sheet");

// Insert data and apply styles.
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;

// Save your Excel document.
workBook.SaveAs("NewExcelFile.xlsx");
```

## Exporting to Various Formats

Files can be saved or exported to popular spreadsheet formats with ease:

```cs
// Fluent saving to multiple formats
workSheet.SaveAs("NewExcelFile.xls");
workSheet.SaveAs("NewExcelFile.xlsx");
workSheet.SaveAsCsv("NewExcelFile.csv");
workSheet.SaveAsJson("NewExcelFile.json");
workSheet.SaveAsXml("NewExcelFile.xml");
```

## Styling Cells and Ranges

Styling of Excel cells and ranges is handled through the `IronXL.Range.Style` object:

```cs
// Define cell value and apply styles.
workSheet["A1"].Value = "Hello World";
workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");
workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;
```

## Sorting Data in Excel

Sorting is straightforward in IronXL, allowing range sorting via the `Range` class:

```cs
using IronXL;

WorkBook workBook = WorkBook.Load("test.xls");
WorkSheet workSheet = workBook.WorkSheets.First();

// Obtain a range.
Range range = workSheet["A2:A8"];

// Sort the range in ascending order.
range.SortAscending();
workBook.Save();
```

## Modifying Formulas

Setting and editing an Excel formula is as straightforward as adding an `=` sign at the beginning:

```cs
// Assign a formula
workSheet["A1"].Value = "=SUM(A2:A10)";

// Retrieve the computed value
decimal sum = workSheet["A1"].DecimalValue;
```

## Why Opt for IronXL?

IronXL offers a developer-friendly API for handling Excel documents in .NET environments efficiently and without the need for Excel Interop installations.

## Further Steps

To delve deeper into IronXL, review our comprehensive [.NET API Reference](https://ironsoftware.com/csharp/excel/object-reference/) styled like MSDN documentation.