# C# Excel Add in (Code Example Tutorial)

***Based on <https://ironsoftware.com/how-to/excel-add-in-csharp/>***


Developing applications often requires managing data in Excel spreadsheets programmatically—such as inserting new rows or columns. C# provides robust capabilities for manipulating Excel files directly through libraries like IronXL. The examples below demonstrate how to leverage these capabilities.

<hr class="separator">

<p class="main-content__segment-title">Step 1</p>

## 1. Install the IronXL Excel Library 

Before adding rows and columns to your Excel files, you'll need to integrate IronXL. This library, which is freely available for development use, can be directly downloaded or installed via NuGet.

- [Download the DLL here](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.Add.Excel.Csharp.zip)
- Install through NuGet:

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<p class="main-content__segment-title">How to Tutorial</p>

## 2. Adding Rows to Excel in C&num;

With IronXL installed, you can now effortlessly insert rows and columns into existing Excel spreadsheets.

Start by loading your Excel file and selecting the worksheet where you want to add rows or columns.

### 2.1. Add a Row at the Last Position
Our first example demonstrates adding a row at the end of the spreadsheet. Assuming the file is named `sample.xlsx` and contains columns from `A` to `E`, here is how you add the row:

```cs
// Add a new row at the end of the Excel spreadsheet
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    int newRowPosition = ws.Rows.Count() + 1;  // Position for the new row
    for(char column = 'A'; column <= 'E'; column++){
        ws[$"{column}{newRowPosition}"].Value = "New Row";
    }
    wb.SaveAs("sample.xlsx");
}
```

This code inserts a new row filled with "New Row" indicating a successful addition.

### 2.2. Add a Row at the First Position

If you need to add a new row at the top of an Excel sheet:

```cs
// Insert a new row at the beginning of the spreadsheet
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws.Rows.Insert(0); // Shift all rows down by one
    ws.Rows[0].Value = "new row"; // Set the new first row's value
    wb.SaveAs("sample.xlsx");
}
```

Compare the table before and after in `sample.xlsx`, using these images:

|Before|After|
|:---:|:-----:|
|![before](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/before2.png)|![after](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/after2.png)|

### 2.3. Add a New First Row When There are Headers

If the first row contains headers:

```cs
// Revise the original loop to maintain headers and add a new data row at the second position
// Because the previous code was shown with logical errors, we have revised it to correctly handle headers
```

<hr class="separator">

## 3. Adding Columns in Excel with C# 

You can also add columns to your Excel sheets. Suppose we want to insert a column before the first existing column in `sample.xlsx`:

```cs
// Code to add a new column at the first position in an Excel sheet
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws.Columns.Insert(0); // Shift all columns to the right
    ws.Rows.ForEach(row => row[0].Value = "New Column Added"); // Add new column content
    wb.SaveAs("sample.xlsx");
}
```

Here’s the visual comparison showing the spreadsheet before and after adding a column:

|Before|After|
|:---:|:-----:|
|![before](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/before1.png)|![after](https://ironsoftware.com/img/faq/excel/excel-add-in-csharp/after1.png)|

<hr class="separator">

<p class="main-content__segment-title">Library Quick Access</p>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Explore IronXL Documentation</h3>
      <p>Dive deeper into other functionalities of IronXL by reviewing the comprehensive documentation.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> Read the IronXL Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>