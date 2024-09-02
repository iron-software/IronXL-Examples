# C# Excel Add-In: A Comprehensive Code Example Tutorial

When you're building applications that interface with Excel files, it might be necessary to programatically manipulate these files — even without opening Excel. For instance, adding rows or columns to an existing spreadsheet could be required. In this tutorial, we'll explore how to do just that by employing the "Excel: Add" capabilities with C#.

<hr class="separator">

<p class="main-content__segment-title">Step 1</p>

## 1. Obtain IronXL Excel Library

To start manipulating Excel files, such as adding rows or columns, you first need to obtain the IronXL Excel Library. You can download this library at no cost for development purposes. Acquire the DLL via direct download [here](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.Add.Excel.Csharp.zip) or opt for the NuGet package installation method available at [NuGet IronXL.Excel](https://www.nuget.org/packages/IronXL.Excel).

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<p class="main-content__segment-title">How to Tutorial</p>

## 2. Add Rows in Excel with C\#

With IronXL installed, let's proceed by adding new rows and columns to existing Excel files. Begin by loading your spreadsheet and identifying the specific worksheet in which to operate.

### 2.1. Insert a Row at the End

Here’s how to append a row at the end of the worksheet named `sample.xlsx` that contains columns `A` to `E`.

```cs
/**
Add Row at the End
anchor-add-row-in-last-position
**/
using IronXL;

static void Main(string[] args)
{
    WorkBook workbook = WorkBook.Load("sample.xlsx");
    WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
    int newRowIdx = sheet.Rows.Count() + 1;
    for (char column = 'A'; column <= 'E'; column++)
    {
        sheet[$"{column}{newRowIdx}"].Value = "New Row";
    }
    workbook.SaveAs("sample.xlsx");
}
```

This code snippet adds a new row at the bottom of "Sheet1" in `sample.xlsx`.

### 2.2. Insert a Row at the Beginning

Adding a new row at the beginning of a spreadsheet is just as easy. The following example demonstrates this process.

```cs
/**
Add Row at the Beginning
anchor-add-row-in-first-position
**/
using IronXL;

static void Main(string [] args)
{
    WorkBook workbook = WorkBook.Load("sample.xlsx");
    WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
    
    sheet.InsertRow(1);
    sheet["A1:E1"].Value = "New Row"; // Set New Row to all columns from A to E

    workbook.SaveAs("sample.xlsx");
}
```

Observe the transformation in the `sample.xlsx` before and after the row insertion.

|Before|After|
|:---:|:-----:|
|![before](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/before2.png)|![after](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/after2.png)|

### 2.3. Add Row at the Top Excluding Column Headers

If the first row contains column headers, add a new row immediately after it:

```cs
// Adding row under headers, assuming first row is headers.
sheet.InsertRow(2); // Inserts a new row at the second position
```

<hr class="separator">

## 3. Adding a Column in C\#

Introducing a new column in an existing worksheet can be needed for data alignment or stitching in additional datasets.

```cs
/**
Add Column
anchor-excel-add-column-in-c-num
**/
using IronXL;

static void Main(string [] args)
{
    WorkBook workbook = WorkBook.Load("sample.xlsx");
    WorkSheet sheet = workbook.GetWorkSheet("Sheet1");
    for (int row = 1; row <= sheet.Rows.Count(); row++)
    {
        for (char col = 'F'; col >= 'B'; col--)
        {
            sheet[$"{col}{row}"].Value = sheet[$"{(char)(col - 1)}{row}"].Value;
        }
        sheet[$"A{row}"].Value = "New Column"; // Add new column data
    }
    workbook.SaveAs("sample.xlsx");
}
```

Preview the adjustments made to `sample.xlsx`:

|Before|After|
|:---:|:-----:|
|![before](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/before1.png)|![after](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/after1.png)|

<hr class="separator">

<p class="main-content__segment-title">Library Quick Access</p>

<div class="tutorial-section">
<div class="row">
<div class="col-sm-8">
<h3>Explore IronXL Documentation</h3>
<p>Dive deeper into IronXL with the comprehensive documentation, offering insight into additional functions and features specific to Excel and C# interaction.</p>
<a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Explore the IronXL Documentation <i class="fa fa-chevron-right"></i></a>
</div>
<div class="col-sm-4">
<div class="tutorial-image">
<img style="max-width: 110px; width: 100px; height: 140px;" alt="" src="https://ironsoftware.com/img/svgs/documentation.svg" class="img-responsive add-shadow" width="100" height="140">
</div>
</div>
</div>
</div>