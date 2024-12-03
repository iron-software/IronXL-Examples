# C# Edit Excel File

***Based on <https://ironsoftware.com/how-to/csharp-edit-excel-file/>***


Editing Excel files in C# demands careful handling to prevent unintended changes which can significantly alter the outputed document. Having access to straightforward and effective code snippets is crucial for minimizing errors while editing or modifying Excel files programmatically. This guide will show you how to efficiently and correctly manipulate Excel files using C# with the help of the IronXL library.

---

<p class="main-content__segment-title">Step 1</p>

## 1. C# Edit Excel Files using the IronXL Library

This tutorial utilizes the IronXL library, a vital C# tool for handling Excel files. To begin, you need to install IronXL in your project (which is free for development purposes).

Feel free to [Download IronXL.zip](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.Edit.Excel.Csharp.zip) or visit the [NuGet package page](https://www.nuget.org/packages/IronXL.Excel) to install it.

After installation, you can proceed as follows:

```shell
Install-Package IronXL.Excel
```

---

<p class="main-content__segment-title">How to Tutorial</p>

## 2. Edit Specific Cell Values

We'll start by modifying specific cell values within an Excel spreadsheet.

First, import the Excel file and select the worksheet you wish to modify. Below is an example of how to perform this task:

```cs
// Import and edit spreadsheet
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx"); // Load an Excel spreadsheet
    WorkSheet ws = wb.GetWorkSheet("Sheet1"); // Select the worksheet
    ws.Rows[3].Columns[1].Value = "New Value"; // Modify a specific cell
    wb.SaveAs("sample.xlsx"); // Save the changes
}
```

Here are snapshots of the Excel sheet `sample.xlsx` before and after the changes:

|Before|After|
|:---:|:-----:|
|![before](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_before1.png)|![after](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_after1.png)|

To modifiy a cell using its address:

```cs
ws["B4"].Value = "New Value"; // Modify a specific cell using its address
```

---

## 3. Edit Full Row Values

It's straightforward to update an entire row in an Excel spreadsheet with a static value:

```cs
// Edit Full Row Values
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws.Rows[3].Value = "New Value";        
    wb.SaveAs("sample.xlsx");
}
```

Here are the before and after snapshots:

|Before|After|
|:---:|:-----:|
|![before](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_before2.png)|![after](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_after2.png)|

To edit a specific range within a row:

```cs
ws["A3:E3"].Value = "New Value";
```

---

## 4. Edit Full Column Values

Similarly, modifying an entire column in an Excel spreadsheet is just as easy:

```cs
// Edit Full Column Values
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws.Columns[1].Value = "New Value";
    wb.SaveAs("sample.xlsx");
}
```

Here's the visual comparison:

|Before|After|
|:---:|:-----:|
|![before](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_before4.png)|![after](https://ironsoftware.com/img/faq/excel/csharp-edit-excel-file/doc5_after4.png)|

---

## 5. Edit Full Row with Dynamic Values

Dynamic values can be applied to specific rows using IronXL, allowing each cell within the row to get a unique value based on its index:

```cs
// Edit row with dynamic values
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    for (int i = 0; i < ws.Columns.Count(); i++)
    {
        ws.Rows[3].Columns[i].Value = "New Value " + i.ToString();
    }
    wb.SaveAs("sample.xlsx");
}
```

Below are the before and after visuals:

|Before|After|
|:---:|:-----:|

The article continues with examples of editing columns dynamically, replacing values in spreadsheets, and methods to remove rows and worksheets. Please reach out to our development team for further details on using IronXL in your projects.

---

<p class="main-content__segment-title">Library Quick Access</p>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>IronXL Library Documentation</h3>
      <p>Explore the extensive features of the IronXL C# Library to edit, style, or delete your Excel workbooks efficiently.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> IronXL Library Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>