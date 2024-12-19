# VB .NET Read & Create Excel Files (Code Example Tutorial)

***Based on <https://ironsoftware.com/how-to/vb-net-excel-files/>***


For developers working in VB .NET, a robust and straightforward method to manage Excel files is crucial. In this tutorial, we'll explore how to utilize IronXL to manipulate Excel files in VB.NET, allowing us to read and create spreadsheets in various formats like `.xls`, `.xlsx`, `.csv`, and `.tsv`. We'll also dive into ways to customize cell styles and populate data programmatically.

<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Excel for VB.NET Library

Begin by integrating the IronXL Excel library for VB.NET into your project. This can be done either by downloading the DLL from [DLL Download](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.vb.net.excel.files.zip) or by using [NuGet](https://www.nuget.org/packages/IronXL.Excel). IronXL will be instrumental in our walkthrough, particularly for rapid Excel data handling in our VB.NET applications.

```shell
Install-Package IronXL.Excel
```

<hr class="separator">

<h4 class="tutorial-segment-title">How To Tutorial</h4>

## 2. Create Excel Files in VB.NET

IronXL offers an efficient way to generate an Excel file using VB.NET. Let's start by creating a workbook and then explore how to populate it with data and apply style settings to the cells.

### 2.1. Create Excel File

First, initialize a new `WorkBook`:

```vb
Dim wb As New WorkBook
```
This initializes a new `.xlsx` format file.

### 2.2. Create XLS File

To create a file with the `.xls` format, use the following code:

```vb
Dim wb As New WorkBook(ExcelFileFormat.XLS)
```

### 2.3. Create Worksheet

Next, we create a worksheet within the workbook:

```vb
Dim ws1 As WorkSheet = wb.CreateWorkSheet("Sheet1")
```

This creates a new worksheet named `Sheet1`.

### 2.4. Create Multiple Worksheets

You can add additional worksheets like this:

```vb
Dim ws2 As WorkSheet = wb.CreateWorkSheet("Sheet2")
Dim ws3 As WorkSheet = wb.CreateWorkSheet("Sheet3")
```

<hr class="separator">

## 3. Insert Data into Worksheet

### 3.1. Insert Data into Cells

You can insert data directly into cells like this:
```vb
ws1("A1").Value = "Hello World"
```
This places the string "Hello World" into cell `A1` of `ws1`.

### 3.2. Insert Data into Range

To insert data across a range of cells:
```vb
ws1("A3:A8").Value = "NewValue"
```
This will populate cells from `A3` to `A8` in `ws1` with the string "NewValue".

### 3.3. Create and Edit Worksheets Example

Here's how you can create an Excel file and edit it:

```vb
' Create and manipulate Excel file
Imports IronXL

Sub Main()
    Dim wb As New WorkBook(ExcelFileFormat.XLSX)
    Dim ws1 As WorkSheet = wb.CreateWorkSheet("Sheet1")
    ws1("A1").Value = "Hello"
    ws1("A2").Value = "World"
    ws1("B1:B8").Value = "RangeValue"
    wb.SaveAs("Sample.xlsx")
End Sub
```
To save this file in a specific path:
``` vb
wb.SaveAs(@"E:\IronXL\Sample.xlsx")
```

Here is the screenshot of our newly created Excel file `sample.xlsx`:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/vb-net-excel-files/doc5-1.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/vb-net-excel-files/doc5-1.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

It is straightforward to create Excel files with `IronXL` in a VB.NET application.

<hr class="separator">

## 4. Read Excel File in VB.NET

IronXL also simplifies reading `.xlsx` files. Here are the steps to load and use data from an Excel file:

### 4.1. Access Excel File in Project

To load an Excel file into your project:
```vb
Dim wb As WorkBook = WorkBook.Load("sample.xlsx")
```

### 4.2. Access Specific WorkSheet

Here's how to select specific worksheets within your workbook:

#### 4.2.1. By Sheet Name

```vb
Dim ws As WorkSheet = wb.GetWorkSheet("sheet1")
```
#### 4.2.2. By Sheet Index

```vb
Dim ws As WorkSheet = wb.WorkSheets(0)
```
#### 4.2.3. Default Sheet

```vb
Dim ws As WorkSheet = wb.DefaultWorkSheet()
```
#### 4.2.4. First Sheet

```vb
Dim ws As WorkSheet = wb.WorkSheets.FirstOrDefault()
```

After acquiring the worksheet, you can fetch and manipulate the data as required.

<hr class="separator">