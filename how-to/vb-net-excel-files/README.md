# VB.NET Guide on Reading & Creating Excel Files

For developers seeking a streamlined and straightforward method to handle Excel files in VB.NET, we offer an informative guide using IronXL. This tutorial covers the basics of reading Excel files with VB.NET, creating spreadsheets across various formats (`.xls`, `.xlsx`, `.csv`, and `.tsv`), and enhancing cell styles and content in VB.NET applications.


<hr class="separator">

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Excel Library for VB.NET

Initiate your journey by integrating the IronXL Excel Library into your VB.NET project. You can download the required DLLs from [DLL Download](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.vb.net.excel.files.zip) or fetch it via NuGet at [NuGet](https://www.nuget.org/packages/IronXL.Excel). This library is essential for efficiently managing Excel data in your VB.NET applications and is free for development purposes.


```shell
Install-Package IronXL.Excel
```

<hr class="separator">

<h4 class="tutorial-segment-title">How To Tutorial</h4>

## 2. Generate Excel Files in VB.NET

IronXL simplifies the process of generating Excel files in a VB.NET environment. With this library, not only can you create files, but also modify cell properties such as fonts and borders.

### 2.1. Initialize a Workbook

To begin, create a new Workbook object:

```vb
Dim wb As New WorkBook
```
This code instantiates a new Excel file with the `.xlsx` extension as the default.

### 2.2. Specify XLS Format

If you need a file with an `.xls` format, you can specify it as follows:

```vb
Dim wb As New WorkBook(ExcelFileFormat.XLS)
```

### 2.3. Add a Worksheet

Next, add a worksheet to your workbook:

```vb
Dim ws1 As WorkSheet = wb.CreateWorkSheet("Sheet1")
```
This line of code creates a new worksheet named `Sheet1` within the workbook `wb`.

### 2.4. Add Multiple Worksheets

You can add multiple worksheets similarly:

```vb
Dim ws2 As WorkSheet = wb.CreateWorkSheet("Sheet2")
Dim ws3 As WorkSheet = wb.CreateWorkSheet("Sheet3")
```

<hr class="separator">

## 3. Populate Data into a Worksheet

### 3.1. Enter Data into Specific Cells

Data can be entered into worksheet cells with ease:

```vb
ws1("A1").Value = "Hello World"
```
The above line inputs "Hello World" into cell `A1` of the worksheet `ws1`.

### 3.2. Populate a Range of Cells

Data can also be populated across a range:

```vb
ws1("A3:A8").Value = "NewValue"
```
This will fill cells `A3` to `A8` in worksheet `ws1` with "NewValue".

### 3.3. Example of Creating and Editing a Workbook

To demonstrate, let’s create a new Excel file named `Sample.xlsx` and populate it with data:

```vb
' Example: Create and Edit Excel
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
**Note:** By default, the new Excel file will be created in the `bin>Debug` folder of your project. To specify another path, use:
``` vb
wb.SaveAs(@"E:\IronXL\Sample.xlsx")
```

Here is a screenshot of the newly created Excel file `sample.xlsx`:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/vb-net-excel-files/doc5-1.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/vb-net-excel-files/doc5-1.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

This outline demonstrates the ease of creating and managing Excel files with `IronXL` in VB.NET applications.

<hr class="separator">

## 4. Reading an Excel File in VB.NET

IronXL also facilitates the reading of Excel (`xlsx`) files. Here’s how you can load and access data from an Excel document in your project.

### 4.1. Load an Excel File

To access a file, instantiate a `WorkBook` with the path to your Excel file:

```vb
WorkBook wb = WorkBook.Load("sample.xlsx") ' Load the Excel file into `wb`
```

### 4.2. Access a Specific Worksheet

There are several methods to select a particular worksheet from a workbook:

#### 4.2.1. By Sheet Name

```vb
Dim ws As WorkSheet = wb.GetWorkSheet("Sheet1") ' Select by sheet name
```

#### 4.2.2. By Sheet Index

```vb
Dim ws As WorkSheet = wb.WorkSheets(0) ' Select the first sheet by index
```

#### 4.2.3. Using Default or First Sheet

```vb
Dim ws As WorkSheet = wb.DefaultWorkSheet() ' Select the default sheet
```
```vb
Dim ws As WorkSheet = wb.WorkSheets.FirstOrDefault() ' Select the first available sheet
```

Once the desired worksheet (`ws`) is obtained, you can fetch and manipulate data from it as needed.

<hr class="separator">

## 5. Retrieving Data from a Worksheet

Data retrieval is straightforward:

```vb
Dim intValue As Integer = sheet("A2").IntValue ' Retrieve an integer value
Dim strValue As String = sheet("A2").ToString() ' Retrieve a string value
```

### 5.1. Extract Data from a Specific Column

You can also loop through cells in a specific column to extract data:

```vb
For Each cell In sheet("A2:A10")
    Console.WriteLine("Value is: {0}", cell.Text)
Next
```

This loop displays values from cell `A2` to `A10`. Here’s a complete example of this process:

```vb
' Load and Retrieve Values
Imports IronXL
Sub Main()
    Dim wb As WorkBook = WorkBook.Load("sample.xlsx")
    Dim ws As WorkSheet = wb.WorkSheets.FirstOrDefault()
    For Each cell In ws("A2:A10")
        Console.WriteLine("Value is: {0}", cell.Text)
    Next
    Console.ReadKey()
End Sub
```
This will output data from the specified cell range.

Here’s a screenshot showing the output in the console and the related Excel file `Sample.xlsx`:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/vb-net-excel-files/doc3-output2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/vb-net-excel-files/doc3-output2.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

Further, you can explore advanced functionalities like aggregate functions (Sum, Min, or Max) on Excel data, which are well explained in the [tutorial on using Excel functions](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/#sample-function-sum).

<hr class="separator">

<h4 class="tutorial-segment-title">Tutorial Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>API Documentation</h3>
      <p>Access detailed API documentation for IronXL, which offers numerous functionalities to enhance Excel processing in your VB.NET projects. Explore features, functions, classes, and more through the comprehensive API guide. </p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">API Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>

This comprehensive guide provides a thorough overview of handling Excel files effectively using IronXL in a VB.NET environment, ensuring that developers can integrate, manipulate, and utilize Excel data seamlessly within their applications.