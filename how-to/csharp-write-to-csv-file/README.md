# How to Write CSV in .NET

***Based on <https://ironsoftware.com/how-to/csharp-write-to-csv-file/>***


Curious about using C# to write to CSV? Discover how IronXL simplifies the process of writing data into CSV files in the .NET framework.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Writing CSV in .NET</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-add-ironxl-to-your-project">Incorporate the IronXL Library</a></li>
        <li><a href="#anchor-2-create-an-excel-workbook">Generate a Workbook in C#</a></li>
        <li><a href="#anchor-3-save-workbook-to-csv">Export Excel Workbook to CSV</a></li>
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

## 1. Incorporate IronXL into Your Project

If you have yet to install IronXL, follow these steps:

* Launch Visual Studio and access the Project menu
* Select Manage NuGet Packages
* Find 'IronXL.Excel' through the search bar
* Choose Install

Alternatively, input this command in the Developer Command Prompt:

```shell
Install-Package IronXL.Excel
```

For additional help, consult our tutorials at [IronXL Guide](https://ironsoftware.com/csharp/excel/docs/).

You can download the sample project [here](https://ironsoftware.com/csharp/excel/downloads/csharp-write-to-csv.zip).

<hr class="separator">

<h4 class="tutorial-segment-title">Tutorial Guide</h4>

## 2. Generate an Excel Workbook

Start by creating a simple Excel workbook with the data outlined below:

<center>
  <div class="center-image-wrapper">
    <a
      href="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/normal-excel-data-to-be-exported-to-csv.png"
      target="_blank"
    >
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/normal-excel-data-to-be-exported-to-csv.png"
        alt="Sample Excel data to be exported to CSV"
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
        Sample Excel data to be exported to CSV
      </span>
    </div>
  </div>
</center>

Then, incorporate the IronXL namespace to enable CSV file writing capabilities in C# using IronXL.

```cs
using IronXL;
```

<hr class="separator">

## 3. Export Workbook to CSV

This code snippet utilizes the `WorkBook` object's `Load` method to load a file into Excel and then uses `SaveAs` to save it as a CSV:

```cs
/**
Save as CSV File
anchor-save-workbook-to-csv
**/
private void button3_Click(object sender, EventArgs e)
{
    WorkBook wb = WorkBook.Load("Normal_Excel_File.xlsx"); // Load an Excel file
    wb.SaveAs("Excel_To_CSV.csv"); // Save it as CSV, naming it as 'Excel_To_CSV.Sheet1.csv'
}
```

Here's how the resulting CSV file appears when opened with a simple Text Editor like Notepad:

<center>
  <div class="center-image-wrapper">
    <a
      href="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/output-csv-file.png"
      target="_blank"
    >
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/output-csv-file.png"
        alt="Displayed CSV file"
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
        Displayed CSV file
      </span>
    </div>
  </div>
</center>

<hr class="separator">

<h4 class="tutorial-segment-title">Library Quick Access</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>IronXL API Reference Documentation</h3>
      <p>Explore further and learn about merging, unmerging, and manipulating cells in Excel sheets through our detailed IronXL API Reference Documentation.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank"> Discover IronXL API Reference Docs <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>