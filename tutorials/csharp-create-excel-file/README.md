# C&num; Excel File Handling Tutorial

This guide describes how to interact with Excel files using C#. Topics covered include data validation, database conversion, web API integrations, and formula editing. The examples herein use the IronXL .NET Excel library.

IronXL enables reading and editing of Microsoft Excel files using C#. It does not necessitate Microsoft Excel or Interop, offering an API that is faster and more user-friendly than `Microsoft.Office.Interop.Excel`.

## Advantages of IronXL include

* Dedicated support from .NET engineers
* Smooth installation through Microsoft Visual Studio
* Free trial for development needs. Licenses start from $749.

IronXL simplifies the process of reading and creating Excel files with C# and VB.NET.

## Working with XLS and XLSX File Formats Using IronXL

The workflow for reading Excel files with IronXL includes the following steps:

1. Install the IronXL Excel Library via the [NuGet package](https://www.nuget.org/packages/IronXL.Excel/).
2. Use `WorkBook.Load `to read XLS, XLSX, or CSV files.
3. Extract cell values with intuitive syntax: `sheet["A11"].DecimalValue`

```csharp
using IronXL;
using System;
using System.Linq;

// XLSX, XLS, CSV, and TSV are all supported formats for reading
WorkBook workBook = WorkBook.Load("test.xlsx");
WorkSheet workSheet = workBook.WorkSheets.First();

// Select cells in Excel notation and return their calculated value
int cellValue = workSheet["A2"].IntValue;

// Reading from a range of cells is simple.
foreach (var cell in workSheet["A2:A10"])
{
    Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
}

// Advanced Operations
// Calculate aggregate values such as minimum, maximum, and sum
decimal sum = workSheet["A2:A10"].Sum();

// The library is LINQ compatible
decimal max = workSheet["A2:A10"].Max(c => c.DecimalValue);
```

Examples in subsequent sections of this tutorial (along with the sample project code) will work on three Excel spreadsheets, illustrated below:

## Step-by-Step Guide

### 1. Download the IronXL C# library for free

Start by installing the `IronXL.Excel` library, which enhances the .NET framework with Excel functionality.

`IronXL.Excel` can be installed using the NuGet package method, or manually by downloading the DLL to your project's global assembly cache.
How to Install the IronXL NuGet Package

1. Right-click the project in Visual Studio and select "Manage NuGet Packages ..."
2. Search for the `IronXL.Excel` package and install it

Alternatively, use the NuGet Package Manager Console:

1. Access the Package Manager Console
2. Type `Install-Package IronXL.Excel`

```shell
PM > Install-Package IronXL.Excel
```

For more information, visit the package on the NuGet website.
Manual Installation

IronXL .NET Excel DLL can also be downloaded and manually installed into Visual Studio.

### 2. Loading an Excel Workbook

The `WorkBook` class signifies an Excel sheet. To open an Excel file with C#, use the `WorkBook.Load` method, indicating the file's path.

```csharp
WorkBook workBook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
```

Each `WorkBook` can contain multiple `WorkSheet` objects. Each one represents a single worksheet in the Excel document. Use `WorkBook.GetWorkSheet` to get a specific Excel worksheet.

```csharp
WorkSheet workSheet = workBook.GetWorkSheet("GDPByCountry");
```

#### Creating New Excel Documents

To make a new Excel document, construct a `WorkBook` object with a valid file type.

```csharp
WorkBook workBook = new WorkBook(ExcelFileFormat.XLSX);
```

Note: ExcelFileFormat.XLS supports older versions of Microsoft Excel (95 and previous).

#### Adding a Worksheet to an Excel Document

An IronXL `WorkBook` contains a collection of one or more `WorkSheets`.

To create a new worksheet, call `WorkBook.CreateWorkSheet` and name the worksheet.

```csharp
WorkSheet workSheet = workBook.GetWorkSheet("GDPByCountry");
```

### 3. Accessing Cell Values

#### Reading and Editing a Single Cell

Access individual spreadsheet cell values by pulling the desired cell from its `WorkSheet`.

```csharp
WorkBook workBook = WorkBook.Load("test.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;
IronXL.Cell cell = workSheet["B1"].First();
```

The `Cell` class in IronXL signifies an individual cell in an Excel spreadsheet. It contains properties and methods that let users access and modify the cell's value directly.

Each `WorkSheet` object maintains an index of Cell objects corresponding to each cell value in an Excel worksheet. We reference the desired cell using standard array indexing syntax.

After referencing a `Cell` object, we can read and write its data:

```csharp
IronXL.Cell cell = workSheet["B1"].First();
string value = cell.StringValue;  
Console.WriteLine(value);

cell.Value = "10.3289";
Console.WriteLine(cell.StringValue);
```

#### Reading and Writing a Range of Cell Values

The Range class symbolizes a two-dimensional array of Cell objects, referring to an actual range of Excel cells. Obtain these ranges through the string indexer on a `WorkSheet` object.

Arguments are coordinates of a cell ("A1" for instance) or a span of cells.
