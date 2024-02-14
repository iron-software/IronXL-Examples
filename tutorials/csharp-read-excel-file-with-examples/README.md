# C# Tutorial: Reading Excel Files

This tutorial explores techniques to read an Excel file in C#, alongside executing common tasks such as data validation, database conversion, API integrations, and formula modification. The lessons leverage the IronXL .NET Excel library.

IronXL is an effective tool for reading and managing Microsoft Excel documents using C#. It provides a quicker and more intuitive API than `Microsoft.Office.Interop.Excel`, and additionally, doesn't necessitate having Microsoft Excel or Interop.

## Features of IronXL include

- Comprehensive product support from our .NET engineers
- Hassle-free installation using Microsoft Visual Studio
- A free trial for development needs. Licenses start from `$liteLicense`.

IronXL simplifies both reading and creating Excel files in C# and VB.NET.

* * *

### Steps to Reading XLS and XLSX Excel files with IronXL

Here's an outline of the process for reading Excel files using IronXL:

1. First, install the IronXL Excel Library. You can achieve this via our [NuGet package](https://www.nuget.org/packages/IronXL.Excel/) or by downloading the [.Net Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip).
2. Next, use the `WorkBook.Load` method to read any XLS, XLSX, or CSV file.
3. Finally, use this straightforward syntax to retrieve cell values: `sheet["A11"].DecimalValue`.

```csharp
    using IronXL;
    using System;
    using System.Linq;
    
    // IronXL supports reading from formats like XLSX, XLS, CSV, and TSV
    WorkBook workBook = WorkBook.Load("test.xlsx");
    WorkSheet workSheet = workBook.WorkSheets.First();
    
    // To select cells using Excel notation and return the calculated value:
    int cellValue = workSheet["A2"].IntValue;
    
    // To elegantly read from a range of cells:
    foreach (var cell in workSheet["A2:A10"])
    {
        Console.WriteLine("Cell {0} has value '{1}'", cell.AddressString, cell.Text);
    }
    
    // For advanced operations like calculating aggregate values like Sum, Min, and Max:
    decimal sum = workSheet["A2:A10"].Sum();
    
    // IronXL is LINQ compatible
    decimal max = workSheet["A2:A10"].Max(c =c.DecimalValue);
```

The code snippets used in further sections of this tutorial (along with the sample project code) will operate on three sample Excel spreadsheets (as illustrated below).

![Sample_Spreadsheets](https://ironsoftware.com/img/tutorials/how-to-read-excel-file-csharp/vs-spreadsheets.png)

* * *

## Detailed Steps

### 1. Download the IronXL C# Library for FREE

Begin by installing the `IronXL.Excel` library to introduce Excel functionality into the .NET framework.

You can install `IronXL.Excel` conveniently with our NuGet package, or manually install the [DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) into your project or into your global assembly cache.

### How to Install the IronXL NuGet Package

1. In Visual Studio, perform a right-click on the project and select "Manage NuGet Packages ...".
2. Search for the IronXL.Excel package and click on the Install button to add it to your project.

    ![NuGet_Installation](https://platform.openai.com/img/tutorials/how-to-read-excel-file-csharp/ef-nuget.png)

Another method is using the NuGet Package Manager Console:

1. Access the Package Manager Console.
2. Type `Install-Package IronXL.Excel`.

    ```shell
    PM Install-Package IronXL.Excel
    ```

You can also [check the package on the NuGet website](https://www.nuget.org/packages/IronXL.Excel/).

### Manual Installation

Alternatively, start by downloading the [IronXL .NET Excel DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) and manually installing it into Visual Studio.

## 2. Load an Excel Workbook

An instance of the `WorkBook` class denotes an Excel sheet. To open an Excel File using C#, you can use the `WorkBook.Load` method and specify the target file's path.

```csharp
WorkBook workBook = WorkBook.Load(@"Spreadsheets\\GDP.xlsx");
```

A `WorkBook` can embed multiple `WorkSheet` instances. Each `WorkSheet` signifies a single Excel worksheet in the document. Use `WorkBook.GetWorkSheet` to access a specific Excel worksheet.

```csharp
WorkSheet workSheet = workBook.GetWorkSheet("GDPByCountry");
```

### Creating New Excel Documents

For creating a new Excel document, construct a `WorkBook` instance with a valid file type.

```csharp
WorkBook workBook = new WorkBook(ExcelFileFormat.XLSX);
```

Note: Employ `ExcelFileFormat.XLS` for compatibility with legacy Microsoft Excel versions (95 and earlier).

### Adding a Worksheet to an Excel Document

As explained earlier, an instance of IronXL's `WorkBook` class can hold a collection of one or more `WorkSheet`s.

![WorkBook_Example](https://platform.openai.com/img/tutorials/how-to-read-excel-file-csharp/work-book.png)

To introduce a new worksheet, call `WorkBook.CreateWorkSheet` and provide the name of the worksheet.

```csharp
WorkSheet workSheet = workBook.GetWorkSheet("GDPByCountry");
```

## 3\. Accessing Cell Values

### Reading and Editing a Single Cell

Values of individual spreadsheet cells can be accessed by extracting the intended cell from its `WorkSheet`.

```csharp
WorkBook workBook = WorkBook.Load("test.xlsx");
WorkSheet workSheet = workBook.DefaultWorkSheet;
IronXL.Cell cell = workSheet["B1"].First();
```

IronXL's `Cell` class represents an individual cell in an Excel sheet. It houses properties and methods permitting users to directly access and modify the cell's value.

Each `WorkSheet` instance manages an index of `Cell` instances, each corresponding to a cell value in an Excel worksheet. In the source code above, we reference the desired cell by its row and column index (cell B1 in this case) using standard array indexing syntax.

Here's how you can read and write data to a spreadsheet cell:

```csharp
IronXL.Cell cell = workSheet["B1"].First();
string value = cell.StringValue;   // Reading the cell value as a string
Console.WriteLine(value);

cell.Value = "10.3289";            // Writing a new value to the cell
Console.WriteLine(cell.StringValue);
```

### Reading and Writing Ranges of Cell Values

The `Range` class represents a 2D array of `Cell` instances. This collection signifies a literal range of Excel cells. You can acquire ranges by leveraging the string indexer on a `WorkSheet` instance.

The argument here can either be the coordinate of a cell (like "A1", as shown earlier) or a span of cells.
