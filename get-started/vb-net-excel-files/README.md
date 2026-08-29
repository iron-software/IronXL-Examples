# VB.NET Techniques for Reading and Creating Excel Documents

> Full guide: [VB.NET Techniques for Reading and Creating Excel Documents](https://ironsoftware.com/csharp/excel/get-started/vb-net-excel-files/)


For developers using VB.NET, accessing and manipulating Excel files can often be essential. This guide introduces IronXL, demonstrating how it can facilitate the creation and reading of Excel files - including various formats like `.xls`, `.xlsx`, `.csv`, and `.tsv`. Additionally, you'll learn to customize cell styles and populate data using VB.NET Excel functions.

---

#### Step 1: Setting Up Your Excel Library

### 1. Acquire IronXL for VB.NET

Start by adding the IronXL library to your project. You can download the DLL directly from [DLL Download](https://ironsoftware.com/csharp/excel/packages/IronXL.zip) or integrate via [NuGet](https://www.nuget.org/packages/IronXL.Excel).

```shell
Install-Package IronXL.Excel
```

---

#### Guide: Working with Excel in VB.NET

### 2. Generating Excel Documents

With IronXL, generating Excel documents in VB.NET becomes straightforward. Here's how you can create files and configure cell properties:

#### 2.1. Initiate an Excel Workbook

Easily create a new workbook in default format (`.xlsx`):

```vb
' Initialize a new Excel workbook in XLSX format
Dim wb As New WorkBook
```

#### 2.2. Creating `.xls` Files

If required, initialize a workbook with the `.xls` format:

```vb
' Craft a new Excel workbook using .xls format
Dim wb As New WorkBook(ExcelFileFormat.XLS)
```

#### 2.3. Generating a Worksheet

To add a worksheet to your workbook:

```vb
' Generate a worksheet named "Sheet1"
Dim ws1 As WorkSheet = wb.CreateWorkSheet("Sheet1")
```

#### 2.4. Adding Multiple Worksheets

Creating several worksheets is done similarly:

```vb
' Foster additional worksheets within the workbook
Dim ws2 As WorkSheet = wb.CreateWorkSheet("Sheet2")
Dim ws3 As WorkSheet = wb.CreateWorkSheet("Sheet3")
```

### 3. Populating Data in Worksheets

#### 3.1. Data Entry into Specific Cells

Populate individual cells with data like this:

```vb
' Assign a value to a particular cell
ws1("A1").Value = "Hello World"
```

#### 3.2. Data Entry across a Range

Here’s how you input data across multiple cells:

```vb
' Populate a range of cells with a single value
ws1("A3:A8").Value = "NewValue"
```

#### 3.3. Illustrative Example of Worksheet Manipulation

Below is an entire workflow to create an Excel file and manipulate data:

```vb
' Reference the IronXL namespace
Imports IronXL

Sub Main()
    ' Set up a workbook with multiple sheets and data
    Dim wb As New WorkBook(ExcelFileFormat.XLSX)
    Dim ws1 As WorkSheet = wb.CreateWorkSheet("Sheet1")
    ws1("A1").Value = "Hello"
    ws1("A2").Value = "World"
    ws1("B1:B8").Value = "RangeValue"
    wb.SaveAs("Sample.xlsx")
End Sub
```

**Note:** By default, the `SaveAs` method targets the `bin\Debug` folder of your project, adjust the path as needed. For instance:

```vb
wb.SaveAs(@"E:\IronXL\Sample.xlsx")
```

Check out the results from the generated file `Sample.xlsx`:

[View the created Excel file](https://ironsoftware.com/img/faq/excel/vb-net-excel-files/doc5-1.png)

The ease of generating and handling Excel documents with `IronXL` in your VB.NET applications is evident through these examples.

---

### 4. Reading Excel Documents 

IronXL simplifies the process of reading Excel files. To use an existing Excel file in your project, follow these steps:

#### 4.1. Loading an Excel File Into the Project

To integrate an existing Excel file:

```vb
' Load an existing Excel file into the application
Dim wb As WorkBook = WorkBook.Load("sample.xlsx")
```

#### 4.2. Access Specific Worksheet

Access specific sheets using either the name or index:

```vb
' Retrieve a worksheet by name or index
Dim ws As WorkSheet = wb.GetWorkSheet("Sheet1")  # By name
Dim ws As WorkSheet = wb.WorkSheets(0)          # By index
``` 

### 5. Retrieving Data from Worksheets

Extracting data from worksheets is straightforward:

```vb
' Fetch integer and string values from specific worksheet cells
Dim int_value As Integer = ws("A2").IntValue
Dim str_value As String = ws("A2").ToString()
```

For column-specific data:

```vb
' Loop through a specific column range and extract values
For Each cell In ws("A2:A10")
    Console.WriteLine("Value is: {0}", cell.Text)
Next cell
```

### 6. Utilizing Functions on Worksheet Data

Applying aggregate functions (Sum, Min, Max) to worksheet data is efficient and straightforward:

```vb
' Apply and display aggregate functions on worksheet data
Dim sum As Decimal = ws("G2:G10").Sum()
Dim min As Decimal = ws("G2:G10").Min()
Dim max As Decimal = ws("G2:G10").Max()

Console.WriteLine("Sum is: {0}", sum)
Console.WriteLine("Min is: {0}", min)
Console.WriteLine("Max is: {0}", max)
```

Learn more about these processes in the comprehensive [Excel reading guide](https://ironsoftware.com/csharp/excel/#read-excel).

---

#### Quick Access to Documentation and API References

Visit the following section to explore the API documentation and discover additional functionalities available within IronXL for managing Excel in your VB.NET endeavors:

[Documentation API Reference](https://ironsoftware.com/csharp/excel/object-reference/api/)

---