# C# Create Excel

In this tutorial, we'll explore how to programmatically create Excel spreadsheets in C#. This will include creating new files, setting styles, and populating data all using C#. By leveraging the proper tools and code pieces, you're able to tailor your Excel sheets to meet specific needs. Let's walk through the process of building C# Excel workbooks for your .NET projects step by step.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>How to Create C# Excel Workbooks</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-create-c-excel-spreadsheets-with-ironxl">Obtain IronXL Library for C#</a></li>
        <li><a href="#anchor-4-insert-cell-data">Programmatically insert data into cells and ranges</a></li>
        <li><a href="#anchor-6-save-excel-file">Save the Excel document to a designated path</a></li>
        <li><a href="#anchor-8-c-num-excel-from-datatable">Input data from DataTable into Excel</a></li>
        <li><a href="#anchor-9-set-excel-workbook-style">Configure text, cell, and page styles in the workbook</a></li>
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

We'll be using IronXL, an efficient C# library designed for handling Excel files. It simplifies reading, writing, and editing Excel files. You can download it free for development. Install the library and continue following this tutorial.

[Download to your project](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.Create.Excel.Csharp.Spreadsheets.zip) or use [NuGet to install into Visual Studio](https://www.nuget.org/packages/IronXL.Excel).

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">Tutorial Guide</h4>

## 2. Create Excel WorkBook in C#

After installing IronXL, begin by creating an Excel workbook. Use the `WorkBook.Create()` method from IronXL to initiate a new workbook.

```cs
WorkBook wb = WorkBook.Create();
```

To specify whether to create `.xlsx` or `.xls` file formats, provide the `ExcelFileFormat` enum to the `WorkBook.Create()` method:

```cs
// Create a WorkBook for .xlsx file format
WorkBook wb = WorkBook.Create(ExcelFileFormat.XLSX);

// Create a WorkBook for .xls file format
WorkBook wb = WorkBook.Create(ExcelFileFormat.XLS);
```

With `wb` initialized, you can now proceed to create WorkSheets.

<hr class="separator">

## 3. Create Excel WorkSheet

To add worksheets, utilize IronXL's `Workbook.CreateWorkSheet()` function, specifying the worksheet name:

```cs
WorkSheet ws = wb.CreateWorkSheet("MySheet");
```

`wb` refers to the workbook we created, and `ws` represents a new worksheet. Here’s how to add multiple worksheets:

```cs
WorkSheet ws1 = wb.CreateWorkSheet("Overview");
WorkSheet ws2 = wb.CreateWorkSheet("Data");
```

<hr class="separator">

## 4. Insert Data into Cells

For data insertion, we make use of IronXL’s cell addressing approach:

```cs
WorkSheet["AnyCell"].Value = "SomeData";
```

<hr class="separator">

## 5. Populate Data Ranges

You can also populate ranges using the same syntax:

```cs
WorkSheet["StartCell:EndCell"].Value = "DataAcrossTheRange";
```

You can learn more on handling [C# Excel ranges](https://ironsoftware.com/csharp/excel/#excel-ranges) from the detailed IronXL documentation.

<hr class="separator">

## 6. Save the Excel Document

To save changes and data in your Excel file, use the following method:

```cs
WorkBook.SaveAs("FullPath/FileName");
```

Discover additional examples on [creating Excel spreadsheets in C#](https://ironsoftware.com/csharp/excel/#create-excel-spreadsheet) from the IronXL website.

<hr class="separator">

## 7. Example: Create, Populate and Save

Here’s a complete example illustrating the creation of a new workbook, adding data, and saving the file:

```cs
using IronXL;

static void Main(string[] args)
{
    // Initialize a new Workbook
    WorkBook wb = WorkBook.Create(ExcelFileFormat.XLSX);
    // Create a Worksheet in the Workbook
    WorkSheet ws = wb.CreateWorkSheet("DemoSheet");
    // Insert data into individual cells
    ws["A1"].Value = "Welcome";
    ws["B1"].Value = "To";
    ws["C1"].Value = "IronXL";
    // Populate a range with the same value
    ws["A2:C2"].Value = "Unified Data";
    // Save the workbook to a file
    wb.SaveAs("demo.xlsx");
    Console.WriteLine("Excel file created successfully.");
}
```

Visualize the output and the structure of the generated workbook named `demo.xlsx`:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-create-excel/1excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-create-excel/1excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## 8. C# Excel from DataTable

Utilizing IronXL, converting a DataTable to an Excel file becomes straightforward. Let's create a DataTable, populate it, and save it as an Excel file:

```cs
using IronXL;

static void Main(string[] args)
{
    DataTable table = new DataTable("Contacts");
    table.Columns.Add("ID");
    table.Columns.Add("Name");
    table.Columns.Add("Phone");

    for(int i = 0; i < 5; i++)
    {
        table.Rows.Add(i, "Name" + i, "Phone" + i);
    }

    WorkBook wb = WorkBook.Create(ExcelFileFormat.XLSX);
    WorkSheet ws = wb.CreateWorkSheet("ContactsSheet");

    int rowIndex = 1;
    foreach(DataRow row in table.Rows)
    {
        ws["A" + rowIndex].Value = row["ID"];
        ws["B" + rowIndex].Value = row["Name"];
        ws["C" + rowIndex].Value = row["Phone"];
        rowIndex++;
    }

    wb.SaveAs("contacts.xlsx");
    Console.WriteLine("Data table exported to Excel successfully.");
}
```

Here’s a snapshot of our resulting Excel file:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-create-excel/3excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-create-excel/3excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

<hr class="separator">

## 9. Customize Workbook Styles with IronXL

Adding styles programmatically is just as easy with IronXL. Here’s how you can define and apply styles to cells and ranges:

```cs
// Set cell text to bold
WorkSheet["SpecificCell"].Style.Font.Bold = true;

// Make the text italic
WorkSheet["SpecificRange"].Style.Font.Italic = true;

// Apply a strikeout style
WorkSheet["AnotherCell"].Style.Font.Strikeout = true;

// Set a dotted border style
WorkSheet["CellRange"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dotted;

// Change the border color
WorkSheet["AnotherRange"].Style.BottomBorder.SetColor("#00adee");
```

For a comprehensive example including setting multiple styles:

```cs
using IronXL;

static void Main(string [] args)
{
    WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
    WorkSheet worksheet = workbook.CreateWorkSheet("StyledSheet");
    worksheet["A1:D1"].Value = "Header";
    worksheet["A2:A10"].Value = "Details";

    // Set styles
    worksheet["A1:C1"].Style.Font.Bold = true;
    worksheet["D1"].Style.Font.Italic = true;
    worksheet["A2:A10"].Style.Font.Strikeout = true;
    worksheet["B2:C10"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Dotted;
    worksheet["B2:C10"].Style.BottomBorder.SetColor("#ff6700");

    workbook.SaveAs("styledWorkbook.xlsx");
    Console.WriteLine("Styling applied successfully.");
}
```

Visualize the styled `sample.xlsx`:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/csharp-create-excel/2excel.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/csharp-create-excel/2excel.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

Discover more about programming with Excel worksheets in the comprehensive tutorial on [creating and modifying Excel files in C#](https://ironsoftware.com/csharp/excel/tutorials/csharp-open-write-excel-file/).

<hr class="separator">

<h4 class="tutorial-segment-title">Quick Access to Tutorial</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Documentation on C# Excel Creation</h3>
      <p>Explore detailed API references and documentation for creating Excel workbooks, managing worksheets, and applying styles programmatically with IronXL.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> Explore C# Excel Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100%; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>