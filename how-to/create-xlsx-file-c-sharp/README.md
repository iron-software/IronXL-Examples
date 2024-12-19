# Develop an XLSX File in C#

***Based on <https://ironsoftware.com/how-to/create-xlsx-file-c-sharp/>***


In today's automated business environment, working with Excel spreadsheets is a common requirement for .NET applications. This includes not only creating new spreadsheets but also populating them programmatically with data. This tutorial delves into how you can generate Excel spreadsheets in various formats such as `.xls`, `.xlsx`, `.csv`, and `.tsv`, apply cell styles, and insert data using C#.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Excel File Creation in C#</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-ironxl-dll">Install IronXL</a></li>
        <li><a href="#anchor-2-create-a-workbook">Create a .XLSX File</a></li>
        <li><a href="#anchor-4-insert-data-into-worksheets">Populate Data into WorkSheets and Multiple Cells</a></li>
        <li><a href="#anchor-6-set-excelmetadata-for-excel-files">Configure Metadata for Excel Files</a></li>
        <li><a href="#anchor-7-set-cell-style">Adjust Font Style, Strikeout, Border Style, and more</a></li>
      </ul>
    </div>
    <div class="col-sm-6">
      <div class="download-card">
        <img style="box-shadow: none; width: 308px; height: 320px;" src="https://ironsoftware.com/img/faq/excel/how-to-work.svg" class="img-responsive learn-how-to-img replaceable-img">
      </div>
    </div>
  </div>
</div>

<hr class="separator">

<h2>Steps to Create an XLSX File in C#</h2>

1. Obtain the library for crafting XLSX files.
2. Initialize a `Workbook` object to generate an Excel document.
3. Select a default `Worksheet`.
4. Populate the default `Worksheet` with data.
5. Persist the Excel document to storage.

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Download IronXL DLL

IronXL streamlines the process of creating Excel files in C# projects. [Download the DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.create.xlsx.zip) or use [NuGet](https://www.nuget.org/packages/IronXL.Excel) for a simple setup to use it freely during development.

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">Let's Continue</h4>

## 2. Initialize a Workbook

This tool not only allows data insertion but also the modification of cell properties such as font styles and borders.

### 2.1 Creating an XLSX File

To start, use the following snippet to create a Workbook; this defaults to generating a `.xlsx` file:

```cs
/**
Create XLSX File
anchor-create-a-workbook
**/
WorkBook wb = WorkBook.Create();
```

### 2.2 Creating an XLS File
To create a `.xls` formatted file, use the following code:

```cs
WorkBook wb = WorkBook.Create(ExcelFileFormat.XLS);
```

<hr class="separator">

## 3. Generating Excel WorkSheets

Once you have your Workbook ready, proceed to create an Excel WorkSheet. The following code will initiate a new WorkSheet `ws1` named `sheet1` within the `wb` Workbook:

```cs
WorkSheet ws1 = wb.CreateWorkSheet("sheet1");
```

### 3.1 Create Multiple WorkSheets

Similarly, you can create multiple WorkSheets:

```cs
/**
Create WorkSheets
anchor-create-an-excel-worksheet
**/
WorkSheet ws2 = wb.CreateWorkSheet("sheet2");
WorkSheet ws3 = wb.CreateWorkSheet("sheet3");
```
<hr class="separator">

## 4. Populate Data in WorkSheets

You can now easily insert data into cells in a WorkSheet.
```cs
 worksheet ["CellAddress"].Value = "MyValue";
```

### 4.1 Populate Specific WorkSheet Data

To specifically insert data into a WorkSheet, use the following example that places `Hello World` in the `A1` cell of WorkSheet `ws1`:
```cs
/**
Insert WorkSheet Data
anchor-insert-data-into-worksheets
**/
ws1 ["A1"].Value = "Hello World";
```

### 4.2 Distribute Data Across Multiple Cells

You can also populate a range of cells simultaneously. This code will insert `NewValue` in cells from `A3` to `A8` in WorkSheet `ws1`:
```cs
ws1 ["A3:A8"].Value = "NewValue";
```

... and the explanation continues through more steps like creating a sample project, setting Excel metadata and cell styles...

### Further Steps and Tutorial

Dive deeper into the structured step-by-step guide for Excel file creation with .NET available at the [Create Excel Files Using C# tutorial.](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/) 

... and additional content about gaining quick access to the API Reference etc...