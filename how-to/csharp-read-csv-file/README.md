# Working with CSV Files in C#

***Based on <https://ironsoftware.com/how-to/csharp-read-csv-file/>***


IronXL provides a straightforward solution for reading CSV files in C#. Whether you're dealing with comma-separated values or another delimiter, you'll find IronXL to be exceptionally helpful. Below, we explore various aspects of reading CSV files using IronXL in .NET environments.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>Handling CSV Files in .NET Frameworks</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-install-the-ironxl-library">Set Up the C# Library for Reading CSV Files (IronXL)</a></li>
        <li><a href="#anchor-2-read-csv-files-programmatically">Programmatically Handle CSV Files in C#</a></li>
        <li><a href="#anchor-2-read-csv-files-programmatically">Define File Format and Delimiter</a></li>
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

## 1. Setting Up IronXL

To employ IronXL for reading or manipulating CSV files in technologies like MVC, ASP.NET, or .NET Core, start by installing the library. Follow this simple guide:

* Navigate to the Project menu in Visual Studio
* Choose Manage NuGet Packages
* Look up IronXL.Excel
* Proceed with the installation

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

Alternatively, download the package directly from Iron Software's website here: <a class="js-modal-open" href="https://ironsoftware.com/csharp/excel/packages/IronXL.zip" data-modal-id="trial-license-after-download">Download IronXL.zip</a>

<hr class="separator">

<h4 class="tutorial-segment-title">Tutorial Explained</h4>

## 2. Programmatic CSV Reading

Once IronXL is set up in your project, you're ready to start processing CSV files!

Add the IronXL namespace to your project:

```cs
using IronXL;
```

Then, implement the following code snippet to read and manipulate a CSV file programmatically:

```cs
/**
Read and process a CSV file
**/
private void ReadCsvButton_Click(object sender, EventArgs e)
{
    WorkBook workbook = WorkBook.LoadCSV("Sample_CSV_File.csv", fileFormat: ExcelFileFormat.XLSX, ListDelimiter: ",");
    WorkSheet ws = workbook.DefaultWorkSheet;
    workbook.SaveAs("Output_Excel_File.xlsx");
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
        alt="A CSV file displayed in Notepad"
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
        A CSV file displayed in Notepad
      </span>
    </div>
  </div>
</center>

In this tutorial, we create a `WorkBook` to read the CSV data, deciding upon the file's format and the delimiter. Following this, a `WorkSheet` is generated where the data is placed, and afterwards, you can save this as a new Excel file using a new file name and format.

<center>
  <div class="center-image-wrapper">
    <a
      href="https://ironsoftware.com/img/faq/excel/csharp-read-csv-file/the-csv-file-opened-in-excel.png"
      target="_blank"
    >
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-read-csv-file/the-csv-file-opened-in-excel.png"
        alt="The CSV file viewed in Excel"
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
        The CSV file viewed in Excel
      </span>
    </div>
  </div>
</center>

<hr class="separator">

<p class="main-content__segment-title">Quick Reference</p>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>IronXL API Reference Guide</h3>
      <p>Explore more and learn about merging, splitting, and manipulating Excel cells using the comprehensive IronXL API Reference Guide.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Access the IronXL API Reference <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>