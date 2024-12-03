# C# Excel Creation

***Based on <https://ironsoftware.com/how-to/csharp-create-excel/>***


In this guide, we explore how to programmatically create Excel files using C#. We will delve into various aspects such as generating new Excel files, applying styles to cells, and inserting data seamlessly using C# code. Let's dive into the step-by-step process required to create Excel workbooks in your .NET projects.

<div class="learn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Steps to Work with C# Excel Workbook</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-create-c-excel-spreadsheets-with-ironxl">Download the IronXL Library for C#</a></li>
        <li><a href="#anchor-4-insert-cell-data">Programmatically insert data into cells and ranges</a></li>
        <li><a href="#anchor-6-save-excel-file">Save the Excel file to a specified path</a></li>
        <li><a href="#anchor-8-c-num-excel-from-datatable">Import data from DataTable into Excel</a></li>
        <li><a href="#anchor-9-set-excel-workbook-style">Apply styling to workbook text, cells, and pages</a></li>
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

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Create C# Excel Spreadsheets with IronXL

We'll begin by using IronXL, a robust C# library that simplifies the process of working with Excel files. It is available at no cost for development environments. Install it and follow this tutorial.

[Download to your project](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.Create.Excel.Csharp.Spreadsheets.zip) or go to [NuGet for installation into Visual Studio](https://www.nuget.org/packages/IronXL.Excel).

<br>

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">Step-by-step Implementation</h4>

## 2. C# Create Excel Workbook

With IronXL installed, we're ready to create an Excel Workbook. Utilize the `WorkBook.Create()` function from IronXL.

```cs
WorkBook wb = WorkBook.Create();
```

This command initializes a new Excel Workbook `wb`. Here's how you specify the file format (.xlsx or .xls):

```cs
// Create a C# Workbook
// Choose file format here
WorkBook wb = WorkBook.Create(ExcelFileFormat.XLSX); // For .xlsx extension
WorkBook wb = WorkBook.Create(ExcelFileFormat.XLS); // For .xls extension
```

With `wb`, we can now proceed to generate WorkSheets.

<hr class="separator">

## 3. C# Create Excel Spreadsheet

To add a WorkSheet, use IronXL's `Workbook.CreateWorkSheet()` method, which needs a string for the WorkSheet's name.

```cs
WorkSheet ws = wb.CreateWorkSheet("ExampleSheet");
```

Here, `wb` is the Workbook and `ws` is a new WorkSheet. You could create multiple sheets as needed:

```cs
// Creating multiple WorkSheets in C#
WorkSheet ws1 = wb.CreateWorkSheet("Sheet1");
WorkSheet ws2 = wb.CreateWorkSheet("Sheet2");
```

<hr class="separator">

## 4. Insert Cell Data

Let's start populating data into our WorkSheet. Here is how you can address individual cells in Excel:

```cs
// Inserting data into a specific cell
ws["A1"].Value = "Data";
```

<hr class="separator">

## 5. Insert Data in Range

To manipulate several cells at once, utilize the `Range` functionality:

```cs
// Insert data across a range of cells
ws["A1:B2"].Value = "New Value";
```

This populates "New Value" across all cells from A1 to B2. Further details on working with ranges can be found at [C# Excel Ranges](https://ironsoftware.com/csharp/excel/#excel-ranges).

<hr class="separator">

## 6. Save Excel File

After data insertion, it's crucial to save the Excel file at the desired location:

```cs
// Saving the Excel Workbook
wb.SaveAs("YourPath/Filename.xlsx");
```

Discover more about creating Excel spreadsheets in C# at [C# Create Excel Spreadsheet examples](https://ironsoftware.com/csharp/excel/#create-excel-spreadsheet).

<hr class="separator">

## 7. Complete Example: Create, Insert Data, and Save

```cs
// Full example of creating and saving an Excel Workbook
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Create(ExcelFileFormat.XLSX);
    WorkSheet ws = wb.CreateWorkSheet("Sheet1");
    // Inserting data into individual cells
    ws["A1"].Value = "Welcome";
    ws["A2"].Value = "To";
    ws["A3"].Value = "IronXL";
    // Inserting data across a range
    ws["C3:C8"].Value = "Multiple Values";
    // Save the Excel file
    wb.SaveAs("demo.xlsx");
    Console.WriteLine("Workbook created successfully.");
    Console.ReadKey();
}
```

Here's a preview of the newly created Excel Workbook named `demo.xlsx`:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-create-excel/1excel.png"><img src="https://ironsoftware.com/img/faq/excel/csharp-create-excel/1excel.png" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## 8. Importing Data from DataTable to Excel

IronXL simplifies the transfer of DataTable data into your Excel file in just a few lines of code:

```cs
// Example of populating an Excel file from a DataTable
using IronXL;
static void Main(string [] args)
{
    // Create new DataTable and populate it with data
    DataTable dt = new DataTable();
    dt.Columns.Add("ID");
    dt.Columns.Add("Name");
    dt.Columns.Add("Phone Number");
    for (int i = 0; i < 5; i++) {
        dt.Rows.Add("ID" + i, "Name" + i, "12345" + i);
    }
    // Create new Excel file
    WorkBook wb = WorkBook.Create(ExcelFileFormat.XLSX);
    WorkSheet ws = wb.CreateWorkSheet("DataSheet");
    // Populate worksheet from DataTable
    int rowIndex = 1;
    foreach (DataRow row in dt.Rows) {
        ws["A" + rowIndex].Value = row["ID"];
        ws["B" + rowIndex].Value = row["Name"];
        ws["C" + rowIndex].Value = row["Phone Number"];
        rowIndex++;
    }
    // Save the Excel file
    wb.SaveAs("dataImport.xlsx");
}
```

Preview of our resulting Excel file:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-create-excel/3excel.png"><img src="https://ironsoftware.com/img/faq/excel/csharp-create-excel/3excel.png" class="img-responsive add-shadow"></a>
	</div>
</center>

## 9. Set Excel Workbook Style

Lastly, let's apply various styles programmatically. IronXL offers extensive options to style individual cells or ranges:

```cs
// Styling individual cells
ws["A1"].Style.Font.Bold = true;
ws["B1"].Style.Font.Italic = true;
ws["C1"].Style.Font.Strikeout = true;
ws["D1"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;
ws["D1"].Style.BottomBorder.SetColor("#0000ff");

// Styling a range of cells
ws["A2:C2"].Style.Font.Bold = true;
ws["D2:F2"].Style.Font.Italic = true;
ws["G2:I2"].Style.Font.Strikeout = true;
ws["J2:L2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dashed;
ws["J2:L2"].Style.BottomBorder.SetColor("#00ff00");
```

Example of application including styling:

```cs
// Comprehensive example showing data insertion and styling
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Create(ExcelFileFormat.XLSX);
    WorkSheet ws = wb.CreateWorkSheet("StyledSheet");
    // Applying styles
    ws["B2:G2"].Value = "Styling Range";          
    ws["B4:G4"].Value = "Another Styling Range";
    
    // Applying various styles
    ws["B2:D2"].Style.Font.Bold = true;
    ws["E2:G2"].Style.Font.Italic = true;
   .ws["B4:D4"].Style.Font.Strikeout = true;
    ws["E4:G4"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dotted;
    ws["E4:G4"].Style.BottomBorder.SetColor("#ff6600");
    
    wb.SaveAs("styledWorkbook.xlsx");
    console.WriteLine("Styling applied successfully.");
    Console.ReadKey();
}
```

Preview of the styled workbook:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-create-excel/2excel.png"><img src="https://ironsoftware.com/img/faq/excel/csharp-create-excel/2excel.png" class="img-responsive add-shadow"></a>
	</div>
</center>

Explore more about working with Excel in C# within the detailed [API Reference for IronXL](https://ironsoftware.com/csharp/excel/object-reference/api/).

<hr class="separator">

<h4 class="tutorial-segment-title">Tutorial Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Documentation for C# Excel Creation</h3>
      <p>Discover extensive documentation on how to create Excel workbooks, manage worksheets, apply cell styles, and much more in the IronXL API Reference.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">C# Create Excel Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>