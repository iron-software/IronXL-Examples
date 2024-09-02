# C# Excel Interop Alternative Using IronXL

Projects often utilize Excel for clear communication, but relying on `Microsoft.Office.Interop.Excel` can lead to the writing of many complex lines of code. In this guide, we will explore using IronXL as an alternative to C# Excel Interop, allowing you to manage Excel files—create, edit, or manipulate—entirely within C# programming environments.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>No Interop Excel with C#</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-download-ironxl-library">Acquire the No Interop Excel Library</a></li>
        <li><a href="#anchor-2-access-excel-file-data">Access Excel Files in C#</a></li>
        <li><a href="#anchor-3-create-new-excel-file">Programmatically Create and Populate a New Excel File</a></li>
        <li><a href="#anchor-4-modify-existing-excel-file">Edit Existing Files: Update Data, Remove Rows, and More</a></li>
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

<h2>Alternate Approach to Excel Interop</h2>

1. Install a library to handle Excel documents.
2. Open the `Workbook` and add the current Excel file.
3. Set the default Worksheet.
4. Retrieve the value from the Excel Workbook.
5. Display and process the value.

<h4 class="tutorial-segment-title">Step 1</h4>

## 1. Acquire IronXL Library

[Download the IronXL Library](https://ironsoftware.com/csharp/excel/packages/IronXL.Package.For.interop.excel.zip) or [install with NuGet](https://www.nuget.org/packages/IronXL.Excel) to start using the free library, then proceed step-by-step through this tutorial on operating Excel without Interop. Licenses are available for production environments.

```shell
Install-Package IronXL.Excel
```

<hr class="separator">
<h4 class="tutorial-segment-title">Technical Guide</h4>

## 2. Access Excel File Data

To improve business applications, effective and flawless access to data from Excel documents is required, along with the capability to manipulate them through code. IronXL enables you to load a Workbook using the `WorkBook.Load()` function, from which you can select a workbook and read data.

See the subsequent example demonstrating data retrieval in a C# project.

```cs
//Code to Access Excel Data
using IronXL;
static void Main(string [] args)
{
    // Loading the Excel file
    WorkBook wb = WorkBook.Load("sample.xlsx");
    // Selecting a worksheet from the workbook
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    // Reading a specific cell
    string cellValue = ws["A5"].Value.ToString();
    Console.WriteLine("Single Cell Value:\n   Cell A5: {0} ", cellValue);
    Console.WriteLine("\nFetching Multiple Cell Values with Loop:\n");
    // Iterating over a range of cells
    foreach (var cell in ws["B2:B10"])
    {
        Console.WriteLine("   Value: {0}", cell.Text);
    }
    Console.ReadKey();
}
```

This example demonstrates how to access and display data from an Excel file:

<center>
	<div class="center-image-wrapper">
		<a rel="nofollow" href="https://ironsoftware.com/img/faq/excel/c-sharp-excel-interop/1output.png" target="_blank"><img src="https://ironsoftware.com/img/faq/excel/c-sharp-excel-interop/1output.png" alt="" class="img-responsive add-shadow"></a>
	</div>
</center>

### Working with DataSet and DataTables

It's possible to interact with Excel files as a dataset or datatable. More guidance can be found in the [IronXL DataSet and DataTables documentation](https://ironsoftware.com/csharp/excel/#excel-sql-dataset).

```cs
// Utilizing DataSet and DataTables
using IronXL;
WorkBook wb = WorkBook.Load("sample.xlsx");
WorkSheet ws = wb.GetWorkSheet("Sheet1");
//Excel file to Dataset
DataSet ds = wb.ToDataSet();
//Worksheet to DataTable
DataTable dt = ws.ToDataTable(true);
```

We now proceed to address the creation of new Excel files in our C# project.

<hr class="separator">

## 3. Programmatic Creation of a New Excel File

Creating a new Excel workbook and populating it with data in a C# project is straightforward with IronXL, which provides the `WorkBook.Create()` function for generating new files.

Here's how you can set up new worksheets and fill them with data:

```cs
// Demonstrating the Creation of a New Excel File
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Create();
    WorkSheet ws = wb.CreateWorkSheet("sheet1");
    ws["A1"].Value = "New Value A1";
    ws["B2"].Value = "New Value B2";
    wb.SaveAs("NewExcelFile.xlsx");
}
```

This script generates an Excel file named `NewExcelFile.xlsx` and inserts values into specified cells. You can extend this for more complex data entry tasks.

<b>Important: Always ensure to save changes to your Excel files as demonstrated in the example above.</b>

Explore detailed instructions on [creating Excel workbooks with IronXL](https://ironsoftware.com/csharp/excel/#create-excel-spreadsheet).

<hr class="separator">

## 4. Editing an Existing Excel File

Modifying existing Excel workbooks and introducing updated data can be done programmatically in C#. Here are several operations you can perform:

### Updating Cell Values

Modifying cell values in existing Excel workbooks is straightforward. Here is an example demonstrating how you update cell data:

```cs
// Updating Cell Values
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws["A3"].Value = "Updated A3";  // Changing the content of cell A3
    wb.SaveAs("sample.xlsx");
}
```

To update multiple cells simultaneously, use a range like so:

```cs
ws["A3:C3"].Value = "Updated Value";
```

This changes the content of cells from `A3` to `C3` to "Updated Value". Learn more about [utilizing workspaces with IronXL](https://ironsoftware.com/csharp/excel/#excel-ranges).

### Replacing Cell Values

IronXL excels in its flexibility to replace old cell content with new one seamlessly across different scopes—entire worksheets, specific rows or columns, or designated ranges.

Here's how to specify and replace values effectively:

```cs
// Replacing Values in a Range
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws["B5:G5"].Replace("Low", "Moderate");  // Replacing values from 'Low' to 'Moderate' in the range B5 to G5
    wb.SaveAs("sample.xlsx");
}
```

Learn more about [editing cell values across various ranges with IronXL](https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/#sample-edit-cell-values-in-range).

### Removing Rows

Sometimes, rows need to be removed from an Excel workbook in the course of application development. Here’s how you can remove rows using IronXL:

```cs
// Example of Removing a Row in Excel
using IronXL;
static void Main(string [] args)
{
    WorkBook wb = WorkBook.Load("sample.xlsx");
    WorkSheet ws = wb.GetWorkSheet("Sheet1");
    ws.Rows[2].Remove();  // Removing the second row from the worksheet
    wb.SaveAs("sample.xlsx");
}
```

In this example, the second row of the workbook is removed.

<hr class="separator">

<h4 class="tutorial-segment-title">Quick Access to Tutorials</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>IronXL Reference</h3>
      <p>Explore the comprehensive API reference for IronXL to learn more about its functions, features, classes, and namespaces.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> IronXL API Reference <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>