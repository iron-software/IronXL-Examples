# Generating XLSX Files in C#

In today's business environment where automation is pervasive, handling Excel Spreadsheets within .NET applications becomes crucial. This includes tasks like generating spreadsheets in various formats (`.xls`, `.xlsx`, `.csv`, `.tsv`), styling cells, and programmatically inserting data. We'll explore these functionalities using C#.

<div class="learn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Excel File Creation with C#</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-ironxl-dll">Install IronXL</a></li>
        <li><a href="#anchor-2-create-a-workbook">Generate a .XLSX File</a></li>
        <li><a href="#anchor-4-insert-data-into-worksheets">Populate Data into Worksheets</a></li>
        <li><a href="#anchor-6-set-excelmetadata-for-excel-files">Configure Excel File Metadata</a></li>
        <li><a href="#anchor-7-set-cell-style">Apply Styles to Cells</a></li>
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

1. Install the necessary Excel library to create XLSX files.
2. Initiate the `Workbook` object to start your Excel file.
3. Select a default `Worksheet`.
4. Populate the default `Worksheet` with data.
5. Persist your Excel file on the disk.

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Download IronXL DLL

IronXL offers a straightforward method for creating Excel (`.xlsx`) files in your C# projects. You can [download the DLL](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.create.xlsx.zip) or use a [NuGet package installation](https://www.nuget.org/packages/IronXL.Excel). It's free for development use.

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">Step-by-Step Guide</h4>

## 2. Create A Workbook

IronXL lets you not only insert data but also customize cell properties such as fonts and borders.

### 2.1 Create .XLSX File

Initializing a `WorkBook` object like below will create a new Excel file with the `.xlsx` extension by default.

```cs
// Create XLSX File
WorkBook wb = WorkBook.Create();
```

### 2.2 Create .XLS File
For files with an `.xls` extension, adjust your `WorkBook` creation as follows:

```cs
WorkBook wb = WorkBook.Create(ExcelFileFormat.XLS);
```

<hr class="separator">

## 3. Create Excel Worksheet

Post `WorkBook` creation, you can start adding Worksheets. Here's how to make a new `WorkSheet` called `sheet1` in the `WorkBook` `wb`.

```cs
WorkSheet ws1 = wb.CreateWorkSheet("sheet1");
```

### 3.1 Create Multiple WorkSheets

Creating multiple Worksheets is accomplished in the same manner:

```cs
//**
Create Additional WorkSheets
**/
WorkSheet ws2 = wb.CreateWorkSheet("sheet2");
WorkSheet ws3 = wb.CreateWorkSheet("sheet3");
```
<hr class="separator">

## 4. Populate Data into Worksheets

Inserting data into Worksheet cells is straightforward.
```cs
worksheet["CellAddress"].Value = "MyValue";
```

### 4.1 Populate Specific Worksheet

To insert data into worksheet `ws1`, use the following code snippet. It writes `Hello World` into cell `A1`.
```cs
//**
Insert Specific WorkSheet Data
**/
ws1["A1"].Value = "Hello World";
```

### 4.2 Populate Multiple Cells

You can also populate multiple cells at once using the range functionality. The following code will populate cells from `A3` to `A8` with `NewValue`.
```cs
ws1["A3:A8"].Value = "NewValue";
```

<hr class="separator">

## 5. Example Project 

For our example project, we'll create an Excel file named `Sample.xlsx` and populate it with data.

```cs
//**
Example Project
**/
using IronXL;
static void Main(string[] args)
{
    WorkBook wb = WorkBook.Create();  
    WorkSheet ws1 = wb.CreateWorkSheet("sheet1");                    
    ws1["A1"].Value = "Hello";           
    ws1["A2"].Value = "World";
    ws1["B1:B8"].Value = "RangeValue";
    wb.SaveAs("Sample.xlsx");
}
```

Note: By default, the new Excel file is created in the `bin>Debug` folder of the project. For a custom path, use:
 ```wb.SaveAs(@"E:\IronXL\Sample.xlsx");```

Here's a screenshot of the Excel file `sample.xlsx` we created:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/create-xlsx-file-c-sharp/doc5-1.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/create-xlsx-file-c-sharp/doc5-1.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

Creating Excel files using IronXL in your C# Application is quite simple and intuitive.

<hr class="separator">

## 6. Configure Excel Metadata

IronXL also allows you to set metadata for Excel files:

```cs
//**
Set Metadata
**/
WorkBook wb = WorkBook.Create();
wb.Metadata.Author = "AuthorName";
wb.Metadata.Title = "TitleValue";
```

<hr class="separator">

## 7. Customize Cell Styles

Styling cells in your Excel Worksheets is effortless with IronXL. Here’s how to apply different cell styles:

### 7.1. Apply Font Style

You can easily enhance a font style like so:

```cs
//**
Apply Font Style
**/
!*\