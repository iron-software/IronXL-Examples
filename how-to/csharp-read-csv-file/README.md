# Reading a CSV File using C&num; and IronXL

IronXL offers a straightforward solution for parsing CSV files in C#. Whether the delimiters are commas or something else, the following examples illustrate the process.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Guide to Processing CSV Files in .NET</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-install-the-ironxl-library">Setting Up IronXL for CSV Reading in C#</a></li>
        <li><a href="#anchor-2-read-csv-files-programmatically">Programmatically Handling CSV Files in C#</a></li>
        <li><a href="#anchor-2-read-csv-files-programmatically">Configuring File Format and Delimiters</a></li>
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

## 1. Setup IronXL Library

To begin using IronXL to process CSV files in either MVC, ASP, or dotnet core applications, the first step involves installation. Here's how to do it:

* Open Visual Studio and navigate to the Project menu
* Click on Manage NuGet Packages
* Search for the `IronXL.Excel` package
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
        alt="IronXL.Excel NuGet Package Image"
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

Alternatively, download it directly from Iron Software here: <a class="js-modal-open" href="https://ironsoftware.com/csharp/excel/packages/IronXL.zip" data-modal-id="trial-license-after-download">IronXL Package</a>

<hr class="separator">

<h4 class="tutorial-segment-title">Step-by-Step Guide</h4>

## 2. Programmatic CSV File Reading

Let’s start coding!

Begin by including the IronXL namespace:

```cs
using IronXL;
```

Next, implement the code to read a CSV file using IronXL in C#:

```cs
/** 
 * Load and convert a CSV file
 * Anchor-read-csv-files-programmatically
 */
private void button2_Click(object sender, EventArgs e)
{
    WorkBook workbook = WorkBook.LoadCSV("your_file.csv", fileFormat: ExcelFileFormat.XLSX, listDelimiter: ",");
    WorkSheet ws = workbook.DefaultWorkSheet;
    workbook.SaveAs("Excel_format.xlsx");
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
        alt="A CSV File Displayed in Notepad"
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
        CSV File Viewed in Notepad
      </span>
    </div>
  </div>
</center>

Using the `LoadCSV` method of the `WorkBook` class, specify the CSV file's name, format, and delimiter (comma in this example). Then, manipulate the contents using `WorkSheet` object, followed by saving the file in a new format and name.

<center>
  <div class="center-image-wrapper">
    <a
      href="https://ironsoftware.com/img/faq/excel/csharp-read-csv-file/the-csv-file-opened-in-excel.png"
      target="_blank"
    >
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-read-csv-file/the-csv-file-opened-in-excel.png"
        alt="The CSV File Opened in Excel"
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
        Viewing the CSV File in Microsoft Excel
      </span>
    </div>
  </div>
</center>

<hr class="separator">

<p class="main-content__segment-title">Library Quick Access</p>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>Explore IronXL's API Documentation</h3>
      <p>Delve deeper and learn how to efficiently use IronXL for tasks such as merging, unmerging, and manipulating Excel cells with the comprehensive API Reference Documentation.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> IronXL API Reference Documentation <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>