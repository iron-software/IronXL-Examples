# How to Write CSV in .NET

***Based on <https://ironsoftware.com/how-to/csharp-write-to-csv-file/>***


Interested in learning how to efficiently write CSV files using C#? IronXL simplifies the process of writing data to CSV format within the .NET framework.

<div class="learnn-how-section">
  <div class="row">
    <div class="col-sm-6">
      <h2>How to Write CSV in .NET</h2>
      <ul class="list-unstyled">
        <li><a href="#anchor-1-add-ironxl-to-your-project">Integrate IronXL into Your Project</a></li>
        <li><a href="#anchor-2-create-an-excel-workbook">Construct a Workbook Using C#</a></li>
        <li><a href="#anchor-3-save-workbook-to-csv">Convert and Save the Workbook as a CSV File</a></li>
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

## 1. Integrate IronXL into Your Project

If you haven't yet added IronXL to your project, follow these simple instructions:

* Launch Visual Studio and access the Project menu
* Select Manage NuGet Packages
* Look for IronXL.Excel
* Press the Install button

Alternatively, execute the following command in the Developer Command Prompt:

```shell
Install-Package IronXL.Excel
```

For more detailed instructions, click on this [tutorial link](https://ironsoftware.com/csharp/excel/docs/).

The project file can be downloaded [here](https://ironsoftware.com/csharp/excel/downloads/csharp-write-to-csv.zip).

<hr class="separator">

<h4 class="tutorial-segment-title">Tutorial Overview</h4>

## 2. Construct a Workbook in C#

Initiate a new project by creating an Excel workbook with the following data:

<center>
  <div class="center-image-wrapper">
    <a
      href="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/normal-excel-data-to-be-exported-to-csv.png"
      target="_blank"
    >
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/normal-excel-data-to-be-exported-to-csv.png"
        alt="Normal Excel data to be exported to CSV"
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
        Normal Excel data to be exported to CSV
      </span>
    </div>
  </div>
</center>

Subsequently, include the IronXL Namespace to facilitate CSV writing in C# with IronXL:

```cs
using IronXL;
```

<hr class="separator">

## 3. Convert and Save the Workbook as a CSV

Use the following code snippet to load a file into an Excel workbook and convert it to CSV format. This feature also appends the worksheet’s name to the file name, providing a handy reference to the data source:

```cs
/**
Convert Workbook to CSV File
anchor-save-workbook-to-csv
**/
private void button3_Click(object sender, EventArgs e)
{
    WorkBook workbook = WorkBook.Load("Sample_Excel_File.xlsx"); // Load the Excel file
    workbook.SaveAs("Converted_CSV_File.csv"); // Save it as a CSV, including the worksheet's name in the file name
}
```

The produced CSV file, when viewed in a text editor like Notepad, appears as follows:

<center>
  <div class="center-image-wrapper">
    <a
      href="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/output-csv-file.png"
      target="_blank"
    >
      <img
        class="img-responsive"
        src="https://ironsoftware.com/img/faq/excel/csharp-write-to-csv-file/output-csv-file.png"
        alt="Output CSV file"
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
        Output CSV file
      </span>
    </div>
  </div>
</center>

<hr class="separator">

<h4 class="tutorial-segment-title">Quick Access to the Library</h4>

<div class="tutorial-section">
  <div class="row">
    <div class="col-sm-8">
      <h3>IronXL API Reference Documentation</h3>
      <p>Enhance your projects by exploring how to merge, unmerge, and manipulate cells in Excel documents using the comprehensive IronXL API Reference Documentation.</p>
      <a class="doc-link" href="https://ironsoftware.com/csharp/excel/object-reference/api/" target="_blank">Explore IronXL API Reference <i class="fa fa-chevron-right"></i></a>
    </div>
    <div class="col-sm-4">
      <div class="tutorial-image">
        <img style="max-width: 110px; width: 100px; height: 140px;" alt="" class="img-responsive add-shadow" src="https://ironsoftware.com/img/svgs/documentation.svg" width="100" height="140">
      </div>
    </div>
  </div>
</div>