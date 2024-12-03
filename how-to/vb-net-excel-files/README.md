# VB .NET: Read & Create Excel Files with IronXL (Detailed Code Guide)

***Based on <https://ironsoftware.com/how-to/vb-net-excel-files/>***


Software developers often seek efficient methods to work with Excel files using VB .NET. This guide demonstrates how to employ IronXL to manage and manipulate Excel data effectively for your applications. You'll learn how to craft and modify spreadsheets in several formats such as `.xls`, `.xlsx`, `.csv`, and `.tsv`. Moreover, we’ll explore how to enhance spreadsheet aesthetics by setting cell properties and adding content via VB.NET coding techniques.

---

### Step 1: Introduce IronXL to Your VB.NET Project

#### Get the IronXL Excel Library for VB.NET
Begin by integrating the IronXL library into your VB.NET project either through the [DLL Download](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.vb.net.excel.files.zip) or via [NuGet](https://www.nuget.org/packages/IronXL.Excel):

```shell
Install-Package IronXL.Excel
```

This first step is your gateway to efficiently accessing Excel data within VB.NET projects, using IronXL for comprehensive Excel manipulation.

---

### How To Create Excel Files in VB.NET with IronXL

#### 2. Constructing Excel Files

IronXL enables effortless creation of Excel files in your projects. Here’s how you can create and manage Excel spreadsheets:

##### 2.1 Create an Excel File

Simply start by creating a new `WorkBook`:

```vb
Dim wb As New WorkBook
```
This snippet initializes a new Excel file primarily in `.xlsx` format.

##### 2.2 Generate an XLS File

If you need a file in the `.xls` format, adjust the code like so:

```vb
Dim wb As New WorkBook(ExcelFileFormat.XLS)
```

##### 2.3 Establish a Worksheet

Here’s how a new worksheet is created in the workbook:

```vb
Dim ws1 As WorkSheet = wb.CreateWorkSheet("Sheet1")
```
This code snippet crafts a new worksheet named `Sheet1`.

##### 2.4 Add Multiple Worksheets

Creating additional sheets follows a similar pattern:

```vb
Dim ws2 As WorkSheet = wb.CreateWorkSheet("Sheet2")
Dim ws3 As WorkSheet = wb.CreateWorkSheet("Sheet3")
```

#### 3. Populate Data into Worksheets

##### 3.1 Fill Individual Cells

You can directly insert data into specific cells:

```vb
ws1("A1").Value = "Hello World"
```
This will populate `Hello World` in the `A1` cell of the `ws1` worksheet.

##### 3.2 Input Range of Data

To fill a range of cells effortlessly:

```vb
ws1("A3:A8").Value = "NewValue"
```
This records `NewValue` from cell `A3` to `A8` in `ws1`.

##### 3.3 Example: Create & Modify a Workbook

Here’s how you can create a new Excel file called `Sample.xlsx` and input some data:

```vb
' Create and Edit Excel Sample
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
**Note:** By default, new Excel files are saved in the `bin>Debug` folder of your project. Use `wb.SaveAs(@"E:\IronXL\Sample.xlsx")` to specify a custom path.

Refer to the screenshot of our newly created Excel file `sample.xlsx`:
	![](https://ironsoftware.com/img/faq/excel/vb-net-excel-files/doc5-1.png)

IronXL simplifies the process of crafting Excel files in a VB.NET Application.

---

#### 4. Reading Excel Files in VB.NET

IronXL also streamlines the process of reading Excel (`.xlsx`) files within your VB.NET projects. Here’s how you can load, access, and manipulate the data:

##### 4.1 Retrieve the Excel File

```vb
Dim wb As WorkBook = WorkBook.Load("sample.xlsx")
```

##### 4.2 Handle Specific Worksheets

###### 4.2.1 By Sheet Name

```vb
Dim ws As WorkSheet = wb.GetWorkSheet("sheet1")
```

###### 4.2.2 By Sheet Index

```vb
Dim ws As WorkSheet = wb.WorkSheets(0)
```

###### 4.2.3 First Sheet as Default

```vb
Dim ws As WorkSheet = wb.DefaultWorkSheet()
```

##### Continued Data Access and Manipulations

After acquiring the worksheet, you can fetch data and apply various manipulations to it.

---

#### 5. Retrieve Data from Worksheets

```vb
Dim intValue As Integer = ws("A2").IntValue
Dim strValue As String = ws("A2").ToString()
```

##### 5.1 Fetch Data from Specific Column Range

Retrieve and display data for a specific column:

```vb
For Each cell In ws("A2:A10")
    Console.WriteLine("value is: {0}", cell.Text)
Next cell
```

This command loops through the range from `A2` to `A10`, printing each cell’s text.

---

#### 6. Apply Functions to Data

Execute aggregate functions such as Sum, Min, and Max:

```vb
Dim sum As Decimal = ws("G2:G10").Sum()
Dim min As Decimal = ws("G2:G10").Min()
Dim max As Decimal = ws("G2:G10").Max()
```

Learn more about manipulating data with [Excel Aggregate Functions](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/#sample-function-sum).

---

### Documentation & API Reference Quick Access

Explore comprehensive documentation on IronXL, which offers various features and functions to efficiently work with Excel in VB.NET:

[![](https://ironsoftware.com/img/svgs/documentation.svg)](https://ironsoftware.com/csharp/excel/object-reference/api/)
Access the full [Documentation API Reference](https://ironsoftware.com/csharp/excel/object-reference/api/).

This guide provides a practical approach to creating and reading Excel files using IronXL in VB.NET, showcasing the simplicity and power of IronXL for software developers.