# Read a CSV File in C&num;

***Based on <https://ironsoftware.com/how-to/csharp-read-csv-file/>***


For reading CSV files in C#, IronXL provides a straightforward solution. The following examples demonstrate how to manage CSV files using various delimiters in your code.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Managing CSV Files in .NET Applications</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-install-the-ironxl-library">Setting Up a C# Library for CSV File Management (IronXL)</a></li>
        <li><a href="#anchor-2-read-csv-files-programmatically">Programmatic Reading of CSV Files in C#</a></li>
        <li><a href="#anchor-2-read-csv-files-programmatically">Configure File Format and Delimiter Settings</a></li>
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

<p class="main-content__segment-title">Step 1</p>

## 1. Install the IronXL Library

To begin using IronXL for reading CSV files in MVC, ASP.NET, or .NET Core, start by installing the library. Follow these instructions:

* Open Visual Studio and go to the Project menu
* Choose Manage NuGet Packages
* Enter IronXL.Excel in the search bar
* Click Install

<center>
  <div class="center-image-wrapper">
    <a
      href="https://ironsoftware.com/img/faq/excel/csharp-read-csv-file/ironxl-excel-nuget-package.png"
      target="_blank"
    >
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-read-csv-file/ironxl-excel-nuget-package.png"
        alt="IronXL.Excel NuGet Package"
      >
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">
        Figure 1
      </span>
      <span class="image-description-text_regular">
        -
      </span>
      <span class="image-description-text_italic">
        IronXL.Excel NuGet Package
      </span>
    </div>
  </div>
</center>

Alternatively, download it directly from Iron Software's website here: [https://ironsoftware.com/csharp/excel/packages/IronXL.zip](https://ironsoftware.com/csharp/excel/packages/IronXL.zip)

<hr class="separator">

<h4 class="tutorial-segment-title">How-to Tutorial</h4>

## 2. Read CSV Files Programmatically 

Let's start with your project!

First, add the IronXL namespace:

```cs
using IronXL;
```

Next, use the following C# code snippet to read a CSV file using IronXL:

```cs
/**
Handle CSV file reading
anchor-read-csv-files-programmatically
**/
private void button2_Click(object sender, EventArgs e)
{
    WorkBook workbook = WorkBook.LoadCSV("Path_to_Read_CSV_Ex.csv", fileFormat: ExcelFileFormat.XLSX, ListDelimiter: ",");
    WorkSheet ws = workbook.GetWorkSheet("Sheet1"); // Assuming the worksheet you want to interact with is named "Sheet1"
    workbook.SaveAs("Converted_Csv_To_Excel.xlsx");
}
```

<center>
  <div class="center-image-wrapper">
    <a
      href="https://ironsoftware.com/img/faq/excel/csharp-read-csv-file/a-csv-file-opened-in-notepad.png"
      target="_blank"
    >
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-read-csv-file/a-csv-file-opened-in-notepad.png"
        alt="A CSV file opened in Notepad"
      >
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">
        Figure 2
      </span>
      <span class="image-description-text_regular">
        -
      </span>
      <span class="image-description-text_italic">
        A CSV file opened in Notepad
      </span>
    </div>
  </div>
</center>

This code initializes a `WorkBook` object and uses the `LoadCSV` method to specify the CSV file, format, and delimiter. It then accesses the default worksheet where the CSV content is loaded, and finally, the content is saved into a new file.

<center>
  <div class="center-image-wrapper">
    <a
      href="https://ironsoftware.com/img/faq/excel/csharp-read-csv-file/the-csv-file-opened-in-excel.png"
      target="_blank"
    >
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-read-csv-file/the-csv-file-opened-in-excel.png"
        alt="The CSV file opened in Excel"
      >
    </a>
    <div class="image-description">
      <span class="image-description-text_strong">
        Figure 3
      </span>
      <span class="image-description-text_regular">
        -
      </span>
      <span class="image-description-text_italic">
        The CSV file opened in Excel
      </span>
    </div>
  </div>
</center>

<hr class="separator">

<p class="main-content__segment-title">Library Quick Access</p>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>IronXL API Reference Documentation</h3>
      <p>Discover more and learn how to efficiently manage and manipulate cells within Excel documents using the comprehensive IronXL API Reference Documentation.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> Explore IronXL API Reference Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>