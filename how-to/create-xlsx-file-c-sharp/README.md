# Creating Excel Files in C&num; with IronXL

***Based on <https://ironsoftware.com/how-to/create-xlsx-file-c-sharp/>***


In today's highly automated business environment, there's often a need to manipulate Excel spreadsheets throughout .NET applications. This guide demonstrates how to create and manipulate Excel spreadsheets in various formats such as `.xls`, `.xlsx`, `.csv`, and `.tsv` using C#. You'll learn how to initiate spreadsheets, adjust cell styles, and integrate data programmatically.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Excel File Creation Using C#</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-ironxl-dll">Get Started with IronXL</a></li>
        <li><a href="#anchor-2-create-a-workbook">Begin with .XLSX Files</a></li>
        <li><a href="#anchor-4-insert-data-into-worksheets">Incorporate Data into Worksheets</a></li>
        <li><a href="#anchor-6-set-excelmetadata-for-excel-files">Embed Metadata in Excel Files</a></li>
        <li><a href="#anchor-7-set-cell-style">Customize Cell Styles</a></li>
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

<h2>Getting Started with an Excel XLSX File in C&num;</h2>

1. Acquire the IronXL library for handling Excel documents.
2. Create a new `Workbook`.
3. Select or create a `Worksheet`.
4. Populate the chosen `Worksheet` with data.
5. Persist the file to storage.

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Acquiring IronXL Library

IronXL simplifies the process of crafting Excel (`.xlsx`) files in C# environments. You can [download the DLL directly](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.create.xlsx.zip) or add it via [NuGet](https://www.nuget.org/packages/IronXL.Excel).

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">How to Tutorial</h4>

## 2. Workbook Creation

Empower your applications with capabilities to modify data and set styling for cells such as fonts and borders.

### 2.1 Create `.XLSX` File

Instantiate a `Workbook` for crafting a new `.xlsx` file:

```cs
// Initialize XLSX file creation
WorkBook wb = WorkBook.Create();
```

### 2.2 Create `.XLS` File

For creating a file with `.xls` extension:

```cs
WorkBook wb = WorkBook.Create(ExcelFileFormat.XLS);
```

<hr class="separator">

## 3. Initiate Excel Worksheet

Construct a `Worksheet` in your preferred file format after setting up a `Workbook`:

```cs
WorkSheet ws1 = wb.CreateWorkSheet("sheet1");
```

### 3.1 Manage Multiple Worksheets

You can generate additional `Worksheet` instances similarly:

```cs
// Create additional worksheets
WorkSheet ws2 = wb.CreateWorkSheet("sheet2");
WorkSheet ws3 = wb.CreateWorkSheet("sheet3");
```
<hr class="separator">

## 4. Data Entry in Worksheets

Simple and effective methods for data input into worksheet cells:
```cs
worksheet["CellAddress"].Value = "MyValue";
```

### 4.1 Insert Data into Specific Worksheet

Assign specific data to worksheet `ws1`, for instance:

```cs
// Inserting data into a specific worksheet cell
ws1["A1"].Value = "Hello World";
```

### 4.2 Apply Data to Multiple Cells

Populate multiple cells within a range using:

```cs
// Setting values in a cell range
ws1["A3:A8"].Value = "NewValue";
```

<hr class="separator">

## 5. Excel File Creation Example Project

Create a fresh Excel file titled `Sample.xlsx` and input some initial data:

```cs
// Quick project setup for Excel file creation
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Create();  
    WorkSheet ws1 = wb.CreateWorkSheet("sheet1");                    
    ws1["A1"].Value = "Hello";           
    ws1["A2"].Value = "World";
    ws1["B1:B8"].Value = "RangeValue";
    wb.SaveAs("Sample.xlsx");
}
```

Note: The new file is by default created in the `bin>Debug` directory. To customize the path, modify to: `wb.SaveAs(@"E:\IronXL\Sample.xlsx");`

Here's what our new `Sample.xlsx` looks like:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/create-xlsx-file-c-sharp/doc5-1.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/create-xlsx-file-c-sharp/doc5-1.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

This example illustrates the ease of generating Excel documents using the IronXL library in C# applications.

<hr class="separator">

## 6. Setting Excel Metadata

IronXL also enables you to embed metadata properties within your Excel documents:

```cs
// Embedding author and title metadata in Excel documents
WorkBook wb = WorkBook.Create();
wb.Metadata.Author = "AuthorName";
wb.Metadata.Title = "TitleValue";
```

<hr class="separator">

## 7. Appropriate Cell Styling

IronXL simplifies cell styling settings, offering a full spectrum of style properties for your worksheets.

### 7.1. Adjusting Font Style

Set the font properties:
```cs
// Configuring font style settings for a cell
WorkSheet["CellAddress"].Style.Font.Bold = true;
WorkSheet["CellAddress"].Style.Font.Italic = true;
```

### 7.2. Adding Strikeout

Apply a strikeout to cell text:

```cs
// Applying strikeout styling
WorkSheet["CellAddress"].Style.Font.Strikeout = true;
```

### 7.3. Configuring Border Styles

Define border styles using IronXL:

```cs
// Setting custom border styling
WorkSheet["CellAddress"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dotted;
```

<hr class="separator">

## 8. Comprehensive Cell Styling Project Example

Illustrate multiple cell styling integrations in a singular project:

```cs
// Complete example for setting multiple cell styles
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Create();                     
    WorkSheet ws = wb.CreateWorkSheet("sheet1");

    ws["A1"].Value = "MyVal";
    ws["B2"].Value = "Hello World";

    ws["A1"].Style.Font.Strikeout = true;

    ws["B2"].Style.Font.Bold = true;
    ws["B2"].Style.Font.Italic = true;

    ws["C3"].Style.TopBorder.Type = IronXL.Styles.BorderType.Double;        
    ws["C3"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dotted;
    ws["C3"].Style.LeftBorder.Type = IronXL.Styles.BorderType.Thick;
    ws["C3"].Style.RightBorder.Type = IronXL.Styles.BorderType.SlantedDashDot;
    ws["C3"].Style.BottomBorder.SetColor("#ff6600");
    ws["C3"].Style.TopBorder.SetColor("#ff6600");
    wb.SaveAs("Sample.xlsx");
}
```

Here's a glimpse of the Excel file `sample.xlsx` crafted:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/create-xlsx-file-c-sharp/doc5-2.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/create-xlsx-file-c-sharp/doc5-2.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## 9. Next Steps and Further Learning

For a deeper exploration and detailed steps for creating Excel files using C#, consider reading through the [Create Excel Files Using C# tutorial](https://ironsoftware.com/csharp/excel/tutorials/create-excel-file-net/).

<hr class="separator">

<h4 class="tutorial-segment-title">Quick Tutorial Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
    <div class="col-sm-8">
      <h3>Explore the API Reference</h3>
      <p>Peruse the extensive Documentation for IronXL, which includes details on all namespaces, features, classes, methods, fields, and enums.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Browse the API Reference <i class="fa fa-chevron-right"></i></a>
      </div>
  </div>
</div>